# -*- coding: utf-8 -*-
from __future__ import division

import os
import re
import math
import traceback
from collections import defaultdict, deque

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, forms, script

uidoc = revit.uidoc
doc = revit.doc
output = script.get_output()

# ------------------------------------------------------------
# ТЕПЛОПОТЕРИ ОВ — ЭТАП 2С
# Расчет по квартирам выбранного этажа из связанной АР-модели.
# Модель Revit не изменяется.
# Отчет выводится в окно pyRevit + создается настоящий .xlsx через Excel COM.
# Шаблон не используется. HTML/.xls не используется.
# ------------------------------------------------------------

FT_TO_M = 0.3048
M_TO_FT = 1.0 / 0.3048
MM_TO_FT = 1.0 / 304.8

# Исходные настройки. Пока закреплены в коде, чтобы не плодить лишние окна.
OUTDOOR_TEMP = -33.0
COMMON_TEMP = 16.0
DEFAULT_ROOM_TEMP = 21.0
BATH_TEMP = 24.0
CORNER_ADD_TEMP = 2.0
CORNER_MIN_DIRECTIONS = 2

K_WALL = 0.27
K_WINDOW = 1.09
K_DOOR = 0.53
KIV_W = 250.0
ORIENT_ADD = 0.10
WINDOW_INFILTRATION = 0.10
DEFAULT_STOREY_HEIGHT_M = 3.20

# Защита от фасадных витражей/ленточного остекления.
# Если окно получилось явно фасадной панелью, а не обычным квартирным окном,
# в расчет берём типовой размер окна, чтобы не раздувать теплопотери по BoundingBox.
NORMAL_WINDOW_WIDTH_M = 1.50
NORMAL_WINDOW_HEIGHT_M = 1.50
MAX_REGULAR_WINDOW_WIDTH_M = 2.60
MAX_REGULAR_WINDOW_HEIGHT_M = 2.20
MAX_REGULAR_WINDOW_AREA_M2 = 4.50
BIG_WINDOW_AS_NORMAL_TOKENS = [
    u'витраж', u'витр', u'панорам', u'ленточ', u'фасад',
    u'curtain', u'storefront', u'остекл'
]
WINDOW_WIDTH_PARAM_NAMES = [
    u'Ширина', u'ширина', u'Width', u'WIDTH', u'B', u'b',
    u'ADSK_Ширина', u'ADSK_Размер_Ширина', u'Размер ширина'
]
WINDOW_HEIGHT_PARAM_NAMES = [
    u'Высота', u'высота', u'Height', u'HEIGHT', u'H', u'h',
    u'ADSK_Высота', u'ADSK_Размер_Высота', u'Размер высота'
]

# Смещение для поиска соседнего помещения за стеной.
PROBE_DISTANCES_M = [0.15, 0.30, 0.60, 1.00, 1.50, 2.20]
# При поиске соседнего помещения нельзя проверять только середину стены:
# на сложной АР-геометрии короткие участки у дверей/шахт иначе ошибочно становятся НС.
ADJACENT_SAMPLE_PARAMS = [0.18, 0.35, 0.50, 0.65, 0.82]
# Если за стеной помещение не найдено, но тип/имя стены явно внутренние,
# не считаем её наружной на -33, а считаем как ВНС к МОП/внутреннему пространству 16 °C.
UNKNOWN_INTERNAL_AS_COMMON = True
UNKNOWN_INTERNAL_ADJACENT_TEXT = u'МОП/внутреннее пространство'
WALL_EXTERIOR_TOKENS = [u'наруж', u'фасад', u'exterior', u'external', u'facade']
WALL_INTERIOR_TOKENS = [u'внутр', u'перегород', u'перег.', u'внс', u'межкварт', u'межкомнат', u'коридор', u'моп', u'гкл', u'газобет']
BOUNDARY_GROUP_TOL_M = 0.18
OPENING_TO_SEG_TOL_M = 1.20
WINDOW_TO_SEG_TOL_M = 1.35
DOOR_TO_SEG_TOL_M = 0.75
OPENING_ROOM_SIDE_PROBE_M = [0.12, 0.25, 0.45, 0.75, 1.10, 1.50]

# Точечная настройка фасадных сегментов.
# Короткие наружные возвраты не удаляются, а прибавляются к ближайшей основной фасадной стене.
EXTERIOR_MAIN_MIN_M = 0.70
EXTERIOR_RETURN_MAX_M = 0.45
EXTERIOR_RETURN_ATTACH_TOL_M = 0.35
CORNER_DIRECTION_MIN_TOTAL_M = 0.25
SHARED_BOUNDARY_LINE_TOL_M = 0.45
SHARED_BOUNDARY_OVERLAP_MIN_M = 0.05
SHORT_SEGMENT_INHERIT_LIMIT_M = 0.75

COMMON_ROOM_TOKENS = [
    u'моп', u'коридор', u'лест', u'лк', u'холл', u'тамбур', u'вестиб',
    u'лифтов', u'лифт', u'коляс', u'общ', u'шлюз', u'пожар', u'эвак',
    u'электрощит', u'щитовая', u'тех', u'кладовая инвентар', u'мусор'
]

APARTMENT_PARAM_CANDIDATES = [
    u'Квартира', u'Номер квартиры', u'№ квартиры', u'ADSK_Номер квартиры',
    u'ADSK_Квартира', u'Номер_квартиры', u'Помещение Квартира',
    u'R_Номер_Квартиры', u'Apartment', u'Flat'
]

KIV_TOKENS = [u'кив', u'кiв', u'kiv', u'клапан приточ', u'приточный клапан', u'кпв', u'ns_кив', u'ns_kiv']

FALSE_DOOR_TOKENS = [
    u'проем', u'проём', u'монтаж', u'отверст', u'ниша', u'ниш',
    u'технол', u'условн', u'заглуш', u'встав', u'перегород',
    u'добор', u'панель', u'модуль', u'газобет', u'пер_вн', u'отд_ст'
]
EXTERIOR_DOOR_TOKENS = [u'балкон', u'лодж', u'наруж', u'улич']


def ustr(x):
    try:
        if x is None:
            return u''
        return unicode(x)
    except:
        try:
            return str(x).decode('utf-8')
        except:
            try:
                return str(x)
            except:
                return u''


def low(x):
    return ustr(x).lower().replace(u'ё', u'е')


def safe_float(x, default=0.0):
    try:
        return float(x)
    except:
        return default


def fmt(x, nd=1):
    try:
        return (u'{0:.' + str(nd) + u'f}').format(float(x))
    except:
        return u''


def fmt0(x):
    try:
        return u'{0:.0f}'.format(float(x))
    except:
        return u''


def eid_int(eid):
    try:
        return int(eid.IntegerValue)
    except:
        try:
            return int(eid.Value)
        except:
            return -1


def elem_id(el):
    try:
        return eid_int(el.Id)
    except:
        return -1


def cat_name(el):
    try:
        if el and el.Category:
            return ustr(el.Category.Name)
    except:
        pass
    return u''


def elem_name(el):
    if el is None:
        return u''
    parts = []
    try:
        p = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
        if p:
            v = p.AsValueString()
            if v:
                parts.append(ustr(v))
    except:
        pass
    try:
        n = el.Name
        if n and n not in parts:
            parts.append(ustr(n))
    except:
        pass
    if parts:
        return u' / '.join(parts)
    return cat_name(el) + u' ' + ustr(elem_id(el))


def param_text(el, names):
    if el is None:
        return u''
    for name in names:
        try:
            p = el.LookupParameter(name)
            if p and p.HasValue:
                try:
                    v = p.AsString()
                    if v:
                        return ustr(v)
                except:
                    pass
                try:
                    v = p.AsValueString()
                    if v:
                        return ustr(v)
                except:
                    pass
                try:
                    if p.StorageType == StorageType.Integer:
                        return ustr(p.AsInteger())
                    if p.StorageType == StorageType.Double:
                        return ustr(p.AsDouble())
                except:
                    pass
        except:
            pass
    return u''


def room_number(room):
    try:
        p = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        if p:
            return ustr(p.AsString() or p.AsValueString())
    except:
        pass
    return param_text(room, [u'Номер', u'Number'])


def room_name(room):
    try:
        p = room.get_Parameter(BuiltInParameter.ROOM_NAME)
        if p:
            return ustr(p.AsString() or p.AsValueString())
    except:
        pass
    try:
        return ustr(room.Name)
    except:
        return u''


def room_label(room):
    n = room_number(room)
    nm = room_name(room)
    if n and nm:
        return u'№{0} {1}'.format(n, nm)
    if n:
        return u'№{0}'.format(n)
    return nm


def room_area_m2(room):
    try:
        a = room.Area
        return a * FT_TO_M * FT_TO_M
    except:
        return 0.0


def room_point(room):
    try:
        if room.Location and room.Location.Point:
            return room.Location.Point
    except:
        pass
    try:
        bb = room.get_BoundingBox(None)
        if bb:
            return XYZ((bb.Min.X + bb.Max.X) / 2.0, (bb.Min.Y + bb.Max.Y) / 2.0, (bb.Min.Z + bb.Max.Z) / 2.0)
    except:
        pass
    return None


def is_common_room(room):
    t = low(room_name(room) + u' ' + room_number(room))
    for token in COMMON_ROOM_TOKENS:
        if token in t:
            return True
    return False


def base_temperature_for_room(room):
    t = low(room_name(room) + u' ' + room_number(room))
    bath_tokens = [u'сануз', u'су', u'ванн', u'душ', u'туалет', u'с/у']
    for token in bath_tokens:
        if token in t:
            return BATH_TEMP, u'по имени помещения: санузел/ванная/душевая/туалет'
    if is_common_room(room):
        return COMMON_TEMP, u'по имени помещения: МОП/коридор/лестничная клетка'
    return DEFAULT_ROOM_TEMP, u'по умолчанию для квартирных помещений'


def is_wall(el):
    try:
        return el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_Walls)
    except:
        return False


def is_window(el):
    try:
        return el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_Windows)
    except:
        return False


def is_door(el):
    try:
        return el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_Doors)
    except:
        return False


def is_false_door_candidate(el):
    # В АР-моделях монтажные/условные проёмы часто лежат в категории Doors.
    # Для теплопотерь их нельзя считать дверями.
    t = low(cat_name(el) + u' ' + elem_name(el))
    for token in FALSE_DOOR_TOKENS:
        if token in t:
            return True
    return False


def is_explicit_exterior_door(el):
    # По умолчанию наружные двери в квартирном расчёте не считаем,
    # чтобы не принять монтажный проём в фасадной стене за дверь.
    # Если позже понадобятся балконные/наружные двери — они попадут сюда по имени.
    t = low(cat_name(el) + u' ' + elem_name(el))
    for token in EXTERIOR_DOOR_TOKENS:
        if token in t:
            return True
    return False


def is_kiv_candidate(el):
    t = low(cat_name(el) + u' ' + elem_name(el))
    for token in KIV_TOKENS:
        if token in t:
            return True
    return False


def html_escape(x):
    s = ustr(x)
    s = s.replace(u'&', u'&amp;')
    s = s.replace(u'<', u'&lt;')
    s = s.replace(u'>', u'&gt;')
    s = s.replace(u'"', u'&quot;')
    return s


def strip_bad_filename_chars(x):
    s = ustr(x)
    s = re.sub(ur'[\\/:*?"<>|]+', u'_', s)
    s = s.strip()
    return s or u'Отчет'


def get_revit_links():
    links = []
    for lnk in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        try:
            ldoc = lnk.GetLinkDocument()
            if ldoc:
                links.append((lnk, ldoc))
        except:
            pass
    return links


def choose_link_and_level():
    links = get_revit_links()
    if not links:
        forms.alert(u'В проекте не найдены загруженные связанные модели Revit.', exitscript=True)

    if len(links) == 1:
        link_inst, link_doc = links[0]
    else:
        opts = []
        by_name = {}
        for inst, ldoc in links:
            name = u'{0}  |  {1}'.format(ustr(inst.Name), ustr(ldoc.Title))
            opts.append(name)
            by_name[name] = (inst, ldoc)
        picked = forms.SelectFromList.show(opts, title=u'Выбери связанную АР-модель', button_name=u'Далее')
        if not picked:
            script.exit()
        link_inst, link_doc = by_name[picked]

    levels = list(FilteredElementCollector(link_doc).OfClass(Level))
    levels.sort(key=lambda x: x.Elevation)
    if not levels:
        forms.alert(u'В связанной модели не найдены уровни.', exitscript=True)

    # Показываем только уровни, на которых есть помещения.
    rooms = get_all_rooms(link_doc)
    used_level_ids = set([eid_int(r.LevelId) for r in rooms])
    level_opts = []
    level_by_opt = {}
    for lvl in levels:
        if eid_int(lvl.Id) not in used_level_ids:
            continue
        opt = u'{0}  ({1:.3f} м)'.format(ustr(lvl.Name), lvl.Elevation * FT_TO_M)
        level_opts.append(opt)
        level_by_opt[opt] = lvl

    if not level_opts:
        for lvl in levels:
            opt = u'{0}  ({1:.3f} м)'.format(ustr(lvl.Name), lvl.Elevation * FT_TO_M)
            level_opts.append(opt)
            level_by_opt[opt] = lvl

    picked_level = forms.SelectFromList.show(level_opts, title=u'Выбери этаж для расчёта теплопотерь', button_name=u'Рассчитать')
    if not picked_level:
        script.exit()
    level = level_by_opt[picked_level]

    return link_inst, link_doc, level, levels


def get_all_rooms(link_doc):
    result = []
    try:
        rooms = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
        for r in rooms:
            try:
                if room_area_m2(r) > 0.01:
                    result.append(r)
            except:
                pass
    except:
        pass
    return result


def get_rooms_on_level(link_doc, level):
    rooms = []
    level_id = eid_int(level.Id)
    for r in get_all_rooms(link_doc):
        try:
            if eid_int(r.LevelId) == level_id:
                rooms.append(r)
        except:
            pass
    rooms.sort(key=lambda r: natural_sort_key(room_number(r) + u' ' + room_name(r)))
    return rooms


def natural_sort_key(s):
    s = low(s)
    parts = re.split(ur'(\d+)', s)
    key = []
    for p in parts:
        if p.isdigit():
            key.append(int(p))
        else:
            key.append(p)
    return key


def next_level_height_m(level, levels):
    sorted_lvls = list(levels)
    sorted_lvls.sort(key=lambda x: x.Elevation)
    elev = level.Elevation
    best = None
    for l in sorted_lvls:
        if l.Elevation > elev + 0.001:
            best = l
            break
    if best:
        h = (best.Elevation - elev) * FT_TO_M
        if h > 0.5:
            return h, u'от уровня "{0}" до уровня "{1}"'.format(ustr(level.Name), ustr(best.Name))
    return DEFAULT_STOREY_HEIGHT_M, u'следующий уровень не найден, принято по умолчанию'


def get_room_by_point(rooms, point, exclude_room=None):
    if point is None:
        return None
    for r in rooms:
        if exclude_room is not None and eid_int(r.Id) == eid_int(exclude_room.Id):
            continue
        pts = [point]
        # BoundarySegment часто лежит на уровне пола. Для Room.IsPointInRoom точка
        # на нижней границе может не попасть в помещение, поэтому дублируем проверку
        # на Z центра проверяемого помещения. Это резко снижает ложные НС у МОП.
        rp = room_point(r)
        if rp is not None:
            try:
                pts.append(XYZ(point.X, point.Y, rp.Z))
            except:
                pass
        for ptest in pts:
            try:
                if r.IsPointInRoom(ptest):
                    return r
            except:
                pass
    return None


def curve_mid(curve):
    try:
        return curve.Evaluate(0.5, True)
    except:
        try:
            return XYZ((curve.GetEndPoint(0).X + curve.GetEndPoint(1).X) / 2.0,
                       (curve.GetEndPoint(0).Y + curve.GetEndPoint(1).Y) / 2.0,
                       (curve.GetEndPoint(0).Z + curve.GetEndPoint(1).Z) / 2.0)
        except:
            return None


def curve_len_m(curve):
    try:
        return curve.Length * FT_TO_M
    except:
        return 0.0


def curve_dir_xy(curve):
    try:
        a = curve.GetEndPoint(0)
        b = curve.GetEndPoint(1)
        dx = b.X - a.X
        dy = b.Y - a.Y
        ln = math.sqrt(dx * dx + dy * dy)
        if ln < 1e-9:
            return None
        return XYZ(dx / ln, dy / ln, 0.0)
    except:
        return None


def normal_xy_from_dir(d):
    if d is None:
        return None
    return XYZ(-d.Y, d.X, 0.0)


def add_xyz(a, b, scale):
    return XYZ(a.X + b.X * scale, a.Y + b.Y * scale, a.Z + b.Z * scale)


def segment_key(room, wall_id, curve):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        pts = sorted([(p0.X, p0.Y, p0.Z), (p1.X, p1.Y, p1.Z)])
        return (eid_int(room.Id), int(wall_id),
                int(round(pts[0][0] * 10000.0)), int(round(pts[0][1] * 10000.0)),
                int(round(pts[1][0] * 10000.0)), int(round(pts[1][1] * 10000.0)))
    except:
        return (eid_int(room.Id), int(wall_id), 0, 0, 0, 0)


def dot_xy(a, b):
    try:
        return a.X * b.X + a.Y * b.Y
    except:
        return 0.0


def same_direction_abs(d1, d2):
    if d1 is None or d2 is None:
        return 0.0
    return abs(dot_xy(d1, d2))


def interval_on_dir(curve, d):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        v0 = p0.X * d.X + p0.Y * d.Y
        v1 = p1.X * d.X + p1.Y * d.Y
        return (min(v0, v1), max(v0, v1))
    except:
        return (0.0, 0.0)


def interval_overlap(a, b):
    lo = max(a[0], b[0])
    hi = min(a[1], b[1])
    return max(0.0, hi - lo)


def line_offset_distance(curve_a, curve_b, d):
    try:
        n = XYZ(-d.Y, d.X, 0.0)
        ma = curve_mid(curve_a)
        mb = curve_mid(curve_b)
        if ma is None or mb is None:
            return 999999.0
        ca = ma.X * n.X + ma.Y * n.Y
        cb = mb.X * n.X + mb.Y * n.Y
        return abs(ca - cb)
    except:
        return 999999.0


def build_shared_boundary_map(rooms_on_level, link_doc):
    """Ищет реальные смежные помещения по общей стене/общей границе.
    Это точечная защита от ложных НС на коротких BoundarySegment.
    """
    records_by_wall = defaultdict(list)
    for room in rooms_on_level:
        loops = get_boundary_segments(room)
        for loop in loops:
            for seg in loop:
                try:
                    wall = link_doc.GetElement(seg.ElementId)
                    if wall is None or not is_wall(wall):
                        continue
                    curve = seg.GetCurve()
                    d = curve_dir_xy(curve)
                    mid = curve_mid(curve)
                    if d is None or mid is None:
                        continue
                    wid = elem_id(wall)
                    rec = {
                        'key': segment_key(room, wid, curve),
                        'room': room,
                        'room_id': eid_int(room.Id),
                        'wall_id': wid,
                        'curve': curve,
                        'dir': d,
                        'mid': mid,
                        'length': curve_len_m(curve)
                    }
                    records_by_wall[wid].append(rec)
                except:
                    pass

    result = {}
    line_tol = SHARED_BOUNDARY_LINE_TOL_M * M_TO_FT
    overlap_min = SHARED_BOUNDARY_OVERLAP_MIN_M * M_TO_FT
    for wid, records in records_by_wall.items():
        if len(records) < 2:
            continue
        for a in records:
            best = None
            best_score = None
            ia = interval_on_dir(a['curve'], a['dir'])
            for b in records:
                if a['room_id'] == b['room_id']:
                    continue
                # Границы одной стены у соседних помещений могут идти в противоположных направлениях.
                if same_direction_abs(a['dir'], b['dir']) < 0.985:
                    continue
                off = line_offset_distance(a['curve'], b['curve'], a['dir'])
                if off > line_tol:
                    continue
                ib = interval_on_dir(b['curve'], a['dir'])
                ov = interval_overlap(ia, ib)
                if ov < overlap_min:
                    continue
                # Чем больше перекрытие и меньше смещение линии, тем надежнее сосед.
                score = (-ov, off)
                if best is None or score < best_score:
                    best = b
                    best_score = score
            if best is not None:
                result[a['key']] = best['room']
    return result


def classify_boundary_segment(room, seg, link_doc, rooms_on_level, shared_boundary_map=None):
    curve = seg.GetCurve()
    mid = curve_mid(curve)
    d = curve_dir_xy(curve)
    n = normal_xy_from_dir(d)
    el = link_doc.GetElement(seg.ElementId)

    # Приоритет 1: реальная общая граница с другим помещением на этой же стене.
    # Это не дает коротким стыковочным кускам ошибочно превращаться в наружные стены.
    try:
        if el is not None and shared_boundary_map is not None:
            skey = segment_key(room, elem_id(el), curve)
            adj_from_boundary = shared_boundary_map.get(skey)
            if adj_from_boundary is not None:
                return {
                    'kind': u'ВНС', 'adjacent': adj_from_boundary, 'wall': el, 'outside_normal': n,
                    'assumed_common': False,
                    'reason': u'смежное помещение найдено по общей границе помещений'
                }
    except:
        pass

    if mid is None or n is None:
        # Если нормаль не определилась, не имеем права автоматически делать -33
        # для явно внутренней стены.
        if UNKNOWN_INTERNAL_AS_COMMON and wall_looks_internal(el):
            return {
                'kind': u'ВНС', 'adjacent': None, 'wall': el, 'outside_normal': None,
                'assumed_common': True, 'adjacent_text': UNKNOWN_INTERNAL_ADJACENT_TEXT,
                'reason': u'нормаль не определена, но тип/имя стены внутренние; принято t=16 °C'
            }
        return {
            'kind': u'НС', 'adjacent': None, 'wall': el, 'outside_normal': None,
            'assumed_common': False,
            'reason': u'не удалось определить нормаль сегмента'
        }

    inside_side = None
    for dm in [0.10, 0.20, 0.35, 0.60]:
        off = dm * M_TO_FT
        p1 = add_xyz(mid, n, off)
        p2 = add_xyz(mid, n, -off)
        in1 = False
        in2 = False
        try:
            in1 = room.IsPointInRoom(p1)
        except:
            pass
        try:
            in2 = room.IsPointInRoom(p2)
        except:
            pass
        if in1 and not in2:
            inside_side = 1
            break
        if in2 and not in1:
            inside_side = -1
            break

    # Если определили внутреннюю сторону, наружу от текущего помещения идём в противоположную.
    if inside_side == 1:
        outside_normal = XYZ(-n.X, -n.Y, 0.0)
    elif inside_side == -1:
        outside_normal = n
    else:
        outside_normal = n

    # Приоритет 2: ищем помещение не только в середине стены, а в нескольких точках участка.
    adjacent = find_adjacent_room_along_segment(room, curve, outside_normal, rooms_on_level)

    # Если сторона была определена ненадёжно, пробуем противоположную сторону.
    # Это спасает внутренние стены МОП, где IsPointInRoom на границе/у двери даёт ложный провал.
    if adjacent is None and inside_side is None:
        opposite_normal = XYZ(-outside_normal.X, -outside_normal.Y, 0.0)
        adjacent2 = find_adjacent_room_along_segment(room, curve, opposite_normal, rooms_on_level)
        if adjacent2 is not None:
            adjacent = adjacent2
            outside_normal = opposite_normal

    if adjacent is not None:
        return {
            'kind': u'ВНС', 'adjacent': adjacent, 'wall': el, 'outside_normal': outside_normal,
            'assumed_common': False,
            'reason': u'за границей найдено смежное помещение по нескольким точкам участка'
        }

    # Приоритет 3: если соседнее помещение не найдено, проверяем саму стену.
    # Отсутствие Room за стеной ещё не означает наружный воздух: в АР часто бывают
    # разрывы помещений МОП, шахты, дверные ниши и короткие служебные сегменты.
    if UNKNOWN_INTERNAL_AS_COMMON and wall_looks_internal(el):
        return {
            'kind': u'ВНС', 'adjacent': None, 'wall': el, 'outside_normal': outside_normal,
            'assumed_common': True, 'adjacent_text': UNKNOWN_INTERNAL_ADJACENT_TEXT,
            'reason': u'смежное помещение не найдено, но тип/имя стены внутренние; принято t=16 °C'
        }

    return {
        'kind': u'НС', 'adjacent': None, 'wall': el, 'outside_normal': outside_normal,
        'assumed_common': False,
        'reason': u'за границей не найдено смежное помещение, стена не распознана как внутренняя'
    }

def direction_label_from_normal(n):
    if n is None:
        return u''
    angle = math.atan2(n.Y, n.X)
    deg = (angle * 180.0 / math.pi + 360.0) % 360.0
    dirs = [
        (0, u'В'), (45, u'СВ'), (90, u'С'), (135, u'СЗ'),
        (180, u'З'), (225, u'ЮЗ'), (270, u'Ю'), (315, u'ЮВ'), (360, u'В')
    ]
    best = dirs[0][1]
    bestd = 999
    for d, lbl in dirs:
        dd = abs(deg - d)
        if dd < bestd:
            bestd = dd
            best = lbl
    return best


def exterior_direction_key(n):
    if n is None:
        return None
    angle = math.atan2(n.Y, n.X)
    # Для угловой добавки объединяем направления крупно: по четвертям/полуосьям.
    step = math.pi / 4.0
    return int(round(angle / step))


def get_boundary_segments(room):
    opt = SpatialElementBoundaryOptions()
    try:
        opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
    except:
        pass
    loops = []
    try:
        arr = room.GetBoundarySegments(opt)
        if arr:
            for loop in arr:
                loops.append(list(loop))
    except:
        pass
    return loops


def wall_type_name(wall):
    if wall is None:
        return u''
    try:
        typ = wall.Document.GetElement(wall.GetTypeId())
        if typ:
            return elem_name(typ)
    except:
        pass
    return elem_name(wall)


def wall_type_element(wall):
    try:
        return wall.Document.GetElement(wall.GetTypeId())
    except:
        return None


def wall_function_text(wall):
    typ = wall_type_element(wall)
    if typ is None:
        return u''
    parts = []
    try:
        parts.append(ustr(typ.Function))
    except:
        pass
    try:
        p = typ.get_Parameter(BuiltInParameter.FUNCTION_PARAM)
        if p and p.HasValue:
            try:
                parts.append(ustr(p.AsValueString()))
            except:
                pass
            try:
                parts.append(ustr(p.AsInteger()))
            except:
                pass
    except:
        pass
    return low(u' '.join(parts))


def wall_name_text(wall):
    return low(wall_type_name(wall) + u' ' + elem_name(wall) + u' ' + wall_function_text(wall))


def text_has_token(text, tokens):
    for token in tokens:
        if token in text:
            return True
    return False


def wall_looks_exterior(wall):
    t = wall_name_text(wall)
    if text_has_token(t, WALL_EXTERIOR_TOKENS):
        return True
    try:
        typ = wall_type_element(wall)
        if typ is not None and typ.Function == WallFunction.Exterior:
            return True
    except:
        pass
    try:
        if re.search(ur'(^|[^а-яa-z])нс([^а-яa-z]|$)', t):
            return True
    except:
        pass
    return False


def wall_looks_internal(wall):
    t = wall_name_text(wall)
    if wall_looks_exterior(wall):
        return False
    if text_has_token(t, WALL_INTERIOR_TOKENS):
        return True
    try:
        typ = wall_type_element(wall)
        if typ is not None and typ.Function == WallFunction.Interior:
            return True
    except:
        pass
    return False


def curve_point_at(curve, param):
    try:
        return curve.Evaluate(float(param), True)
    except:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            t = float(param)
            return XYZ(p0.X + (p1.X - p0.X) * t, p0.Y + (p1.Y - p0.Y) * t, p0.Z + (p1.Z - p0.Z) * t)
        except:
            return None


def find_adjacent_room_along_segment(room, curve, normal, rooms_on_level):
    if curve is None or normal is None:
        return None
    for param in ADJACENT_SAMPLE_PARAMS:
        base = curve_point_at(curve, param)
        if base is None:
            continue
        for dm in PROBE_DISTANCES_M:
            p = add_xyz(base, normal, dm * M_TO_FT)
            adjacent = get_room_by_point(rooms_on_level, p, room)
            if adjacent is not None:
                return adjacent
    return None


def line_group_key(curve, kind, adjacent_label, delta_t, k, wall):
    d = curve_dir_xy(curve)
    mid = curve_mid(curve)
    if d is None or mid is None:
        return (kind, adjacent_label, round(delta_t, 2), elem_id(wall))

    # Линия в плане: направление приводим к неориентированному виду.
    dx = d.X
    dy = d.Y
    if dx < -0.0001 or (abs(dx) < 0.0001 and dy < 0):
        dx = -dx
        dy = -dy
    # Нормальная координата линии.
    nx = -dy
    ny = dx
    c = mid.X * nx + mid.Y * ny
    tol = BOUNDARY_GROUP_TOL_M * M_TO_FT
    angle = math.atan2(dy, dx)
    return (kind, adjacent_label, round(delta_t, 2), round(k, 4), elem_id(wall), int(round(angle / 0.08)), int(round(c / tol)))


def line_group_key_no_wall(curve, kind, adjacent_label, delta_t, k):
    d = curve_dir_xy(curve)
    mid = curve_mid(curve)
    if d is None or mid is None:
        return (kind, adjacent_label, round(delta_t, 2), round(k, 4), 'no_wall')

    dx = d.X
    dy = d.Y
    if dx < -0.0001 or (abs(dx) < 0.0001 and dy < 0):
        dx = -dx
        dy = -dy

    nx = -dy
    ny = dx
    c = mid.X * nx + mid.Y * ny
    tol = BOUNDARY_GROUP_TOL_M * M_TO_FT
    angle = math.atan2(dy, dx)
    # Для наружных стен убираем ElementId стены из ключа, чтобы разрезанные фасадные стены
    # одной линии объединялись в одну расчетную строку, как в ручной таблице.
    return (kind, adjacent_label, round(delta_t, 2), round(k, 4), int(round(angle / 0.08)), int(round(c / tol)))


def point_distance_xy_m(p1, p2):
    try:
        dx = p1.X - p2.X
        dy = p1.Y - p2.Y
        return math.sqrt(dx * dx + dy * dy) * FT_TO_M
    except:
        return 999999.0


def curve_endpoints(curve):
    try:
        return [curve.GetEndPoint(0), curve.GetEndPoint(1)]
    except:
        return []


def min_endpoint_distance_m(curve_a, curve_b):
    pts_a = curve_endpoints(curve_a)
    pts_b = curve_endpoints(curve_b)
    if not pts_a or not pts_b:
        return 999999.0
    dmin = 999999.0
    for pa in pts_a:
        for pb in pts_b:
            d = point_distance_xy_m(pa, pb)
            if d < dmin:
                dmin = d
    return dmin


def attach_short_exterior_returns(boundary_items):
    """Короткие наружные куски у примыканий не выбрасываем.
    Если кусок НС короткий и касается основной наружной стены этого же помещения,
    он прибавляется к основной фасадной строке, а не выводится отдельной стеной.
    Это ближе к ручному расчету, где такие возвраты обычно входят в общую фасадную длину.
    """
    targets = []
    for it in boundary_items:
        if it.get('kind') != u'НС':
            continue
        if (it.get('length_m') or 0.0) >= EXTERIOR_MAIN_MIN_M:
            targets.append(it)

    if not targets:
        return

    for it in boundary_items:
        if it.get('kind') != u'НС':
            continue
        if (it.get('length_m') or 0.0) > EXTERIOR_RETURN_MAX_M:
            continue

        best = None
        best_dist = 999999.0
        for tg in targets:
            if tg is it:
                continue
            d = min_endpoint_distance_m(it.get('curve'), tg.get('curve'))
            if d < best_dist:
                best_dist = d
                best = tg

        if best is not None and best_dist <= EXTERIOR_RETURN_ATTACH_TOL_M:
            it['merge_to_item'] = best
            it['orient'] = best.get('orient') or it.get('orient') or u''


def effective_exterior_normal(item):
    tg = item.get('merge_to_item')
    if tg is not None:
        return tg.get('outside_normal')
    return item.get('outside_normal')


def count_exterior_directions_after_merge(boundary_items):
    # Угловую добавку считаем по исходным наружным направлениям, а не по merge_to_item.
    # Иначе маленькая, но реальная боковая наружная стенка у окна теряет своё направление
    # после присоединения к основной фасадной строке.
    totals = defaultdict(float)
    for it in boundary_items:
        if it.get('kind') != u'НС':
            continue
        n = it.get('outside_normal')
        dk = exterior_direction_key(n)
        if dk is None:
            continue
        totals[dk] += it.get('length_m') or 0.0
    return len([dk for dk in totals if totals[dk] >= CORNER_DIRECTION_MIN_TOTAL_M])


def wall_location_curve(wall):
    try:
        loc = wall.Location
        if loc and hasattr(loc, 'Curve'):
            return loc.Curve
    except:
        pass
    return None


def wall_width_m(wall):
    try:
        w = wall.Width * FT_TO_M
        if w > 0:
            return w
    except:
        pass
    try:
        typ = wall.Document.GetElement(wall.GetTypeId())
        if typ:
            for pname in [u'Ширина', u'Width']:
                p = typ.LookupParameter(pname)
                if p and p.HasValue and p.StorageType == StorageType.Double:
                    w = p.AsDouble() * FT_TO_M
                    if w > 0:
                        return w
    except:
        pass
    return 0.0


def wall_axis_interval_info(boundary_curve, wall):
    """Возвращает проекцию участка границы помещения на ось стены.
    Классификация НС/ВНС остается по Room Boundary, но расчетная длина a берется по оси стены.
    """
    try:
        axis = wall_location_curve(wall)
        if axis is None:
            return None
        d = curve_dir_xy(axis)
        if d is None:
            return None
        b0 = boundary_curve.GetEndPoint(0)
        b1 = boundary_curve.GetEndPoint(1)
        t0 = b0.X * d.X + b0.Y * d.Y
        t1 = b1.X * d.X + b1.Y * d.Y
        lo = min(t0, t1)
        hi = max(t0, t1)
        if hi - lo <= 0.0001:
            return None
        wall_iv = interval_on_dir(axis, d)
        return {
            'wall_id': elem_id(wall),
            'lo': lo,
            'hi': hi,
            'wall_lo': wall_iv[0],
            'wall_hi': wall_iv[1],
            'width_m': wall_width_m(wall)
        }
    except:
        return None


def centerline_length_from_axis_infos(axis_infos, fallback_m, allow_unclamped_expand=False):
    """Считает расчетную длину группы.
    Для наружных стен можно не зажимать расширение фактической LocationCurve: в АР
    LocationCurve часто лежит по внутренней/чистовой линии, а ручной расчет берёт
    наружный габарит стены с добором толщины примыкающих ограждений на углах.
    """
    if not axis_infos:
        return fallback_m
    by_wall = defaultdict(list)
    for info in axis_infos:
        if info:
            by_wall[info.get('wall_id')].append(info)
    if not by_wall:
        return fallback_m

    total_ft = 0.0
    for wid, infos in by_wall.items():
        try:
            lo = min([i['lo'] for i in infos])
            hi = max([i['hi'] for i in infos])
            wall_lo = min([i['wall_lo'] for i in infos])
            wall_hi = max([i['wall_hi'] for i in infos])
            width_m = max([safe_float(i.get('width_m'), 0.0) for i in infos])
            expand_ft = max(0.0, width_m * M_TO_FT * 0.5)
            if allow_unclamped_expand:
                lo = lo - expand_ft
                hi = hi + expand_ft
            else:
                lo = max(wall_lo, lo - expand_ft)
                hi = min(wall_hi, hi + expand_ft)
            if hi > lo:
                total_ft += (hi - lo)
        except:
            pass

    if total_ft <= 0:
        return fallback_m

    result_m = total_ft * FT_TO_M
    # Защита от явно неверной LocationCurve: осевая длина не должна внезапно стать в разы больше
    # старой длины границы помещения.
    if fallback_m > 0.01 and result_m > fallback_m * 1.8 + 0.50:
        return fallback_m
    return result_m


def parse_length_m_from_text(value):
    s = low(value).strip()
    if not s:
        return 0.0
    s = s.replace(u' ', u'')
    m = re.search(ur'(\d+(?:[\.,]\d+)?)', s)
    if not m:
        return 0.0
    v = safe_float(m.group(1).replace(u',', u'.'), 0.0)
    if v <= 0:
        return 0.0
    if u'мм' in s or v > 20.0:
        return v / 1000.0
    return v


def parameter_length_m(el, names, bip_names=None):
    if el is None:
        return 0.0
    elems = [el]
    try:
        typ = el.Document.GetElement(el.GetTypeId())
        if typ is not None:
            elems.append(typ)
    except:
        pass

    for e in elems:
        for name in names:
            try:
                p = e.LookupParameter(name)
                if p and p.HasValue:
                    if p.StorageType == StorageType.Double:
                        v = p.AsDouble() * FT_TO_M
                        if v > 0:
                            return v
                    try:
                        txt = p.AsValueString() or p.AsString()
                    except:
                        txt = None
                    v = parse_length_m_from_text(txt)
                    if v > 0:
                        return v
            except:
                pass
        if bip_names:
            for bip_name in bip_names:
                try:
                    bip = getattr(BuiltInParameter, bip_name)
                    p = e.get_Parameter(bip)
                    if p and p.HasValue:
                        if p.StorageType == StorageType.Double:
                            v = p.AsDouble() * FT_TO_M
                            if v > 0:
                                return v
                        try:
                            txt = p.AsValueString() or p.AsString()
                        except:
                            txt = None
                        v = parse_length_m_from_text(txt)
                        if v > 0:
                            return v
                except:
                    pass
    return 0.0


def parse_opening_size_from_name(name):
    s = low(name)
    # 1900x1650, 1900х1650, 1900*1650, 1900×1650h
    m = re.search(ur'(\d{3,4})\s*[xх\*×]\s*(\d{3,4})', s)
    if m:
        a = float(m.group(1)) / 1000.0
        b = float(m.group(2)) / 1000.0
        # Обычно первая величина — ширина, вторая — высота.
        return a, b, u'размер из имени типа'
    return None, None, u''


def opening_size_from_parameters_m(el):
    width = parameter_length_m(el, WINDOW_WIDTH_PARAM_NAMES, [
        'FAMILY_WIDTH_PARAM', 'GENERIC_WIDTH', 'CASEWORK_WIDTH'
    ])
    height = parameter_length_m(el, WINDOW_HEIGHT_PARAM_NAMES, [
        'FAMILY_HEIGHT_PARAM', 'GENERIC_HEIGHT', 'CASEWORK_HEIGHT'
    ])
    if width > 0.01 and height > 0.01:
        return width, height, u'размер из параметров семейства'
    return 0.0, 0.0, u''


def bbox_opening_size_m(el):
    try:
        bb = el.get_BoundingBox(None)
        if bb:
            sx = abs(bb.Max.X - bb.Min.X) * FT_TO_M
            sy = abs(bb.Max.Y - bb.Min.Y) * FT_TO_M
            sz = abs(bb.Max.Z - bb.Min.Z) * FT_TO_M
            a = max(sx, sy)
            b = sz
            if a > 0 and b > 0:
                return a, b, u'размер из BoundingBox семейства'
    except:
        pass
    return 0.0, 0.0, u'размер не найден'


def is_big_window_like_vitrage(el, a, b, src):
    if el is None or not is_window(el):
        return False
    t = low(cat_name(el) + u' ' + elem_name(el))
    has_token = False
    for token in BIG_WINDOW_AS_NORMAL_TOKENS:
        if token in t:
            has_token = True
            break
    too_big = (a > MAX_REGULAR_WINDOW_WIDTH_M or
               b > MAX_REGULAR_WINDOW_HEIGHT_M or
               (a * b) > MAX_REGULAR_WINDOW_AREA_M2)
    # Самый опасный источник — BoundingBox: у фасадных семейств он часто равен панели/витражу,
    # а не расчетному квартирному окну.
    if too_big and (has_token or u'BoundingBox' in ustr(src)):
        return True
    return False


def opening_size_m(el):
    # Последовательность важна: имя/параметры надежнее BoundingBox.
    n = elem_name(el)
    a, b, src = parse_opening_size_from_name(n)
    if not (a and b):
        a, b, src = opening_size_from_parameters_m(el)
    if not (a and b):
        a, b, src = bbox_opening_size_m(el)

    if a and b and is_big_window_like_vitrage(el, a, b, src):
        note = u'витраж/крупное остекление принято как обычное окно {0}×{1} м; исходно {2}×{3} м ({4})'.format(
            fmt(NORMAL_WINDOW_WIDTH_M, 2), fmt(NORMAL_WINDOW_HEIGHT_M, 2), fmt(a, 2), fmt(b, 2), src
        )
        return NORMAL_WINDOW_WIDTH_M, NORMAL_WINDOW_HEIGHT_M, note

    return a, b, src


def element_center(el):
    try:
        if el.Location:
            if hasattr(el.Location, 'Point') and el.Location.Point:
                return el.Location.Point
    except:
        pass
    try:
        bb = el.get_BoundingBox(None)
        if bb:
            return XYZ((bb.Min.X + bb.Max.X) / 2.0, (bb.Min.Y + bb.Max.Y) / 2.0, (bb.Min.Z + bb.Max.Z) / 2.0)
    except:
        pass
    return None


def point_to_curve_distance_xy(p, curve):
    if p is None or curve is None:
        return 999999.0
    try:
        a = curve.GetEndPoint(0)
        b = curve.GetEndPoint(1)
        px = p.X
        py = p.Y
        ax = a.X
        ay = a.Y
        bx = b.X
        by = b.Y
        vx = bx - ax
        vy = by - ay
        wx = px - ax
        wy = py - ay
        vv = vx * vx + vy * vy
        if vv < 1e-12:
            dx = px - ax
            dy = py - ay
            return math.sqrt(dx * dx + dy * dy)
        t = (wx * vx + wy * vy) / vv
        if t < 0:
            t = 0
        if t > 1:
            t = 1
        cx = ax + t * vx
        cy = ay + t * vy
        dx = px - cx
        dy = py - cy
        return math.sqrt(dx * dx + dy * dy)
    except:
        return 999999.0



def closest_point_on_curve_xy(p, curve):
    if p is None or curve is None:
        return None
    try:
        a = curve.GetEndPoint(0)
        b = curve.GetEndPoint(1)
        px = p.X
        py = p.Y
        ax = a.X
        ay = a.Y
        bx = b.X
        by = b.Y
        vx = bx - ax
        vy = by - ay
        wx = px - ax
        wy = py - ay
        vv = vx * vx + vy * vy
        if vv < 1e-12:
            return XYZ(ax, ay, p.Z)
        t = (wx * vx + wy * vy) / vv
        if t < 0:
            t = 0
        if t > 1:
            t = 1
        return XYZ(ax + t * vx, ay + t * vy, p.Z)
    except:
        return None


def point_projection_raw_t_xy(p, curve):
    if p is None or curve is None:
        return None
    try:
        a = curve.GetEndPoint(0)
        b = curve.GetEndPoint(1)
        vx = b.X - a.X
        vy = b.Y - a.Y
        vv = vx * vx + vy * vy
        if vv < 1e-12:
            return None
        wx = p.X - a.X
        wy = p.Y - a.Y
        return (wx * vx + wy * vy) / vv
    except:
        return None


def point_projects_to_curve_xy(p, curve, tolerance_m):
    t = point_projection_raw_t_xy(p, curve)
    if t is None:
        return True
    try:
        length_m = curve_len_m(curve)
        if length_m <= 0.01:
            return True
        tol_ratio = tolerance_m / length_m
        if tol_ratio > 0.35:
            tol_ratio = 0.35
        return (-tol_ratio <= t <= 1.0 + tol_ratio)
    except:
        return True


def safe_point_in_room(room, p):
    try:
        return bool(room.IsPointInRoom(p))
    except:
        return False


def get_host_id(el):
    try:
        if el.Host:
            return eid_int(el.Host.Id)
    except:
        pass
    try:
        return eid_int(el.HostElementId)
    except:
        pass
    return -1


def point_inside_room_towards(room, point):
    """Проверяет, относится ли проём к помещению: от центра проёма идём в сторону
    центра помещения. Центр двери/окна обычно лежит в стене, поэтому простой
    IsPointInRoom по центру часто не работает.
    """
    if point is None:
        return False
    rp = room_point(room)
    if rp is None:
        return False
    try:
        dx = rp.X - point.X
        dy = rp.Y - point.Y
        ln = math.sqrt(dx * dx + dy * dy)
        if ln < 1e-9:
            return False
        ux = dx / ln
        uy = dy / ln
        for dm in OPENING_ROOM_SIDE_PROBE_M:
            p = XYZ(point.X + ux * dm * M_TO_FT, point.Y + uy * dm * M_TO_FT, point.Z)
            try:
                if room.IsPointInRoom(p):
                    return True
            except:
                pass
    except:
        pass
    return False



def opening_belongs_to_room_by_boundary(room, point, item):
    """Проверка принадлежности проёма именно этому помещению.
    Не используем направление на центр помещения как основной критерий: для длинных витражей
    и сложных комнат такой луч легко перескакивает в соседнее помещение. Надежнее брать нормаль
    расчетной границы помещения и проверять точки с внутренней стороны этой границы.
    """
    if point is None or item is None:
        return False
    n = item.get('outside_normal')
    curve = item.get('curve')
    if n is not None:
        for dm in OPENING_ROOM_SIDE_PROBE_M:
            p = XYZ(point.X - n.X * dm * M_TO_FT, point.Y - n.Y * dm * M_TO_FT, point.Z)
            if safe_point_in_room(room, p):
                return True
        cp = closest_point_on_curve_xy(point, curve)
        if cp is not None:
            for dm in OPENING_ROOM_SIDE_PROBE_M:
                p = XYZ(cp.X - n.X * dm * M_TO_FT, cp.Y - n.Y * dm * M_TO_FT, point.Z)
                if safe_point_in_room(room, p):
                    return True
    # Старый способ оставлен только как fallback.
    return point_inside_room_towards(room, point)


def boundary_item_opening_distance_m(center, item):
    try:
        return point_to_curve_distance_xy(center, item.get('curve')) * FT_TO_M
    except:
        return 999999.0


def opening_boundary_allowed(op, item):
    if is_door(op):
        # Для квартир считаем только дверь в МОП/коридор/лестницу.
        # Дверь/проём в наружной стене не считаем автоматически, чтобы не ловить
        # монтажные проёмы и куски стен в категории Doors.
        if item.get('kind') != u'ВНС':
            return False
        adj = item.get('adjacent')
        if adj is None or not is_common_room(adj):
            return False
        return True
    else:
        # Окна и витражи только на наружной границе помещения.
        return item.get('kind') == u'НС'


def nearest_opening_boundary_item(op, center, wall_items, boundary_items):
    hid = get_host_id(op)
    candidates = []
    # Сначала кандидаты по HostId — самый надежный вариант.
    if hid in wall_items:
        candidates.extend(wall_items.get(hid, []))
    # Потом fallback: любые подходящие границы помещения. Это нужно для окон,
    # у которых HostId не совпал с BoundarySegment стены из-за составной/разбитой АР-геометрии.
    for it in boundary_items:
        if it not in candidates:
            candidates.append(it)

    if is_door(op):
        max_dist = DOOR_TO_SEG_TOL_M
    else:
        max_dist = WINDOW_TO_SEG_TOL_M

    best_item = None
    best_score = None
    for item in candidates:
        if not opening_boundary_allowed(op, item):
            continue
        dist_m = boundary_item_opening_distance_m(center, item)
        if dist_m > max_dist:
            continue
        # Центр проёма должен проектироваться на расчетный участок границы,
        # иначе крупный витраж/окно с соседнего помещения может притянуться к ближайшему торцу стены.
        if not point_projects_to_curve_xy(center, item.get('curve'), 0.30):
            continue
        host_bonus = 0 if (hid >= 0 and elem_id(item.get('wall')) == hid) else 1
        # При равной дистанции отдаём приоритет совпадению HostId.
        score = (host_bonus, dist_m)
        if best_item is None or score < best_score:
            best_item = item
            best_score = score
    return best_item


def collect_openings_for_room(room, link_doc, boundary_segments, used_opening_ids):
    # Проёмы назначаются не просто по HostId стены, а по стороне помещения.
    # Это исправляет две ошибки:
    # 1) двери в коридор не должны считаться на -33;
    # 2) окна не должны пропадать, если HostId не совпал с BoundarySegment стены.
    wall_items = defaultdict(list)
    host_wall_ids = set()
    all_boundary_items = []
    for item in boundary_segments:
        wall = item.get('wall')
        curve = item.get('curve')
        if wall is not None and is_wall(wall) and curve is not None:
            wid = elem_id(wall)
            host_wall_ids.add(wid)
            wall_items[wid].append(item)
            all_boundary_items.append(item)

    if not all_boundary_items:
        return []

    openings = []
    cats = [BuiltInCategory.OST_Windows, BuiltInCategory.OST_Doors]
    for bic in cats:
        try:
            elems = FilteredElementCollector(link_doc).OfCategory(bic).WhereElementIsNotElementType()
        except:
            continue
        for op in elems:
            oid = elem_id(op)
            if oid in used_opening_ids:
                continue
            if is_door(op) and is_false_door_candidate(op):
                continue

            c = element_center(op)
            if c is None:
                continue
            best_item = nearest_opening_boundary_item(op, c, wall_items, all_boundary_items)
            if best_item is None:
                continue
            # Проём должен быть с внутренней стороны именно этой расчетной границы помещения.
            if not opening_belongs_to_room_by_boundary(room, c, best_item):
                continue

            openings.append({
                'element': op,
                'boundary_item': best_item,
                'distance_m': boundary_item_opening_distance_m(c, best_item)
            })
            used_opening_ids.add(oid)
    return openings

def collect_current_model_kivs_for_room(room, link_inst):
    result = []
    try:
        inv = link_inst.GetTotalTransform().Inverse
    except:
        inv = None

    cats = [
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_PlumbingFixtures
    ]
    seen = set()
    for bic in cats:
        try:
            elems = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        except:
            continue
        for el in elems:
            if elem_id(el) in seen:
                continue
            if not is_kiv_candidate(el):
                continue
            p = element_center(el)
            if p is None:
                continue
            try:
                lp = inv.OfPoint(p) if inv is not None else p
                if room.IsPointInRoom(lp):
                    result.append(el)
                    seen.add(elem_id(el))
            except:
                pass
    return result


def get_apartment_param_value(room):
    for p in APARTMENT_PARAM_CANDIDATES:
        val = param_text(room, [p])
        if val:
            return val
    return u''


def build_room_adjacency(rooms_on_level, link_doc, shared_boundary_map=None):
    room_by_id = dict([(eid_int(r.Id), r) for r in rooms_on_level])
    adjacency = defaultdict(set)
    for room in rooms_on_level:
        loops = get_boundary_segments(room)
        for loop in loops:
            for seg in loop:
                try:
                    cls = classify_boundary_segment(room, seg, link_doc, rooms_on_level, shared_boundary_map)
                    adj = cls.get('adjacent')
                    if adj is not None:
                        a = eid_int(room.Id)
                        b = eid_int(adj.Id)
                        if a != b:
                            adjacency[a].add(b)
                            adjacency[b].add(a)
                except:
                    pass
    return adjacency, room_by_id


def determine_apartments(rooms_on_level, link_doc, shared_boundary_map=None):
    apartment_rooms = []
    common_rooms = []
    for r in rooms_on_level:
        if is_common_room(r):
            common_rooms.append(r)
        else:
            apartment_rooms.append(r)

    # Если в помещениях уже есть номер квартиры, используем его.
    groups_by_param = defaultdict(list)
    have_param_count = 0
    for r in apartment_rooms:
        ap = get_apartment_param_value(r)
        if ap:
            groups_by_param[ap].append(r)
            have_param_count += 1
    if have_param_count >= max(1, int(len(apartment_rooms) * 0.5)):
        groups = []
        for name in sorted(groups_by_param.keys(), key=natural_sort_key):
            groups.append({'name': u'Квартира ' + name, 'rooms': groups_by_param[name], 'source': u'параметр помещения'})
        return groups, common_rooms, u'по параметрам помещений'

    # Иначе определяем квартиры как связные компоненты квартирных помещений.
    adjacency, room_by_id = build_room_adjacency(rooms_on_level, link_doc, shared_boundary_map)
    apt_ids = set([eid_int(r.Id) for r in apartment_rooms])
    visited = set()
    comps = []
    for rid in sorted(list(apt_ids)):
        if rid in visited:
            continue
        q = deque([rid])
        visited.add(rid)
        comp = []
        while q:
            cur = q.popleft()
            comp.append(cur)
            for nb in adjacency.get(cur, []):
                if nb not in apt_ids:
                    continue
                if nb in visited:
                    continue
                visited.add(nb)
                q.append(nb)
        comps.append(comp)

    # Сортировка квартир по положению на плане.
    comp_items = []
    for comp in comps:
        pts = []
        rooms = []
        for rid in comp:
            r = room_by_id.get(rid)
            if r:
                rooms.append(r)
                p = room_point(r)
                if p:
                    pts.append(p)
        if pts:
            cx = sum([p.X for p in pts]) / len(pts)
            cy = sum([p.Y for p in pts]) / len(pts)
        else:
            cx = cy = 0
        comp_items.append((cy, cx, rooms))
    comp_items.sort(key=lambda x: (-x[0], x[1]))

    groups = []
    idx = 1
    for cy, cx, rooms in comp_items:
        rooms.sort(key=lambda r: natural_sort_key(room_number(r) + u' ' + room_name(r)))
        groups.append({'name': u'Квартира {0}'.format(idx), 'rooms': rooms, 'source': u'авто: связная группа квартирных помещений'})
        idx += 1

    return groups, common_rooms, u'по связности помещений, без явного параметра квартиры'


def make_calc_row(room, ogr, source, element_id, adjacent_text, a_m, b_m, k, dt, orient, note, formula_source, inf=0.0, orient_add=ORIENT_ADD, t_adj=None):
    f = a_m * b_m
    q = f * k * dt
    dob = 1.0 + orient_add + inf
    q_total = q * dob
    return {
        'room_id': eid_int(room.Id),
        'room_number': room_number(room),
        'room_name': room_name(room),
        'room_label': room_label(room),
        'ogr': ogr,
        'source': source,
        'element_id': element_id,
        'adjacent': adjacent_text,
        't_adj': t_adj,
        'a_m': a_m,
        'b_m': b_m,
        'f_m2': f,
        '_gross_f_m2': f,
        '_opening_area_m2': 0.0,
        'k': k,
        'dt': dt,
        'q_w': q,
        'orient_add': orient_add,
        'inf': inf,
        'dob': dob,
        'q_total': q_total,
        'orient': orient,
        'note': note,
        'formula': u'F={0}×{1}={2}; Q={2}×{3}×{4}={5}; Доб.=1+{6}+{7}={8}; Qобщ={5}×{8}={9}'.format(
            fmt(a_m, 3), fmt(b_m, 3), fmt(f, 3), fmt(k, 2), fmt(dt, 1), fmt(q, 1), fmt(orient_add, 2), fmt(inf, 2), fmt(dob, 2), fmt(q_total, 1)
        ),
        'formula_source': formula_source
    }



def recalc_row_after_area_change(row, formula_prefix=None):
    f = safe_float(row.get('f_m2'), 0.0)
    k = safe_float(row.get('k'), 0.0)
    dt = safe_float(row.get('dt'), 0.0)
    q = f * k * dt
    dob = 1.0 + safe_float(row.get('orient_add'), 0.0) + safe_float(row.get('inf'), 0.0)
    q_total = q * dob
    row['q_w'] = q
    row['dob'] = dob
    row['q_total'] = q_total
    if formula_prefix:
        row['formula'] = formula_prefix + u'; Q={0}×{1}×{2}={3}; Доб.=1+{4}+{5}={6}; Qобщ={3}×{6}={7}'.format(
            fmt(f, 3), fmt(k, 2), fmt(dt, 1), fmt(q, 1),
            fmt(row.get('orient_add'), 2), fmt(row.get('inf'), 2), fmt(dob, 2), fmt(q_total, 1)
        )


def subtract_opening_area_from_host_wall(host_item, opening_area_m2, wall_row_by_group_key):
    if host_item is None or opening_area_m2 <= 0:
        return
    key = host_item.get('group_key')
    if key is None:
        return
    wall_row = wall_row_by_group_key.get(key)
    if wall_row is None:
        return
    gross = safe_float(wall_row.get('_gross_f_m2'), wall_row.get('a_m', 0.0) * wall_row.get('b_m', 0.0))
    removed = safe_float(wall_row.get('_opening_area_m2'), 0.0) + opening_area_m2
    wall_row['_opening_area_m2'] = removed
    net = gross - removed
    if net < 0:
        net = 0.0
    wall_row['f_m2'] = net
    note = wall_row.get('note', u'')
    if u'Площадь проёмов вычтена' not in note:
        wall_row['note'] = note + u'; Площадь проёмов вычтена из площади стены'
    formula_prefix = u'F={0}×{1}−проёмы {2}={3}'.format(
        fmt(wall_row.get('a_m'), 3), fmt(wall_row.get('b_m'), 3), fmt(removed, 3), fmt(net, 3)
    )
    recalc_row_after_area_change(wall_row, formula_prefix)


def calculate_room(room, link_doc, link_inst, rooms_on_level, height_m, shared_boundary_map=None):
    base_t, t_source = base_temperature_for_room(room)

    boundary_items = []
    skipped_internal = []
    raw_wall_count = 0
    exterior_dirs = set()
    used_wall_seg_count = 0

    loops = get_boundary_segments(room)
    for loop in loops:
        for seg in loop:
            el = link_doc.GetElement(seg.ElementId)
            if el is None or not is_wall(el):
                continue
            raw_wall_count += 1
            curve = seg.GetCurve()
            cls = classify_boundary_segment(room, seg, link_doc, rooms_on_level, shared_boundary_map)
            kind = cls.get('kind')
            adj = cls.get('adjacent')
            n = cls.get('outside_normal')
            orient = direction_label_from_normal(n)
            boundary_items.append({
                'curve': curve,
                'wall': el,
                'kind': kind,
                'adjacent': adj,
                'outside_normal': n,
                'orient': orient,
                'length_m': curve_len_m(curve),
                'assumed_common': cls.get('assumed_common', False),
                'adjacent_text': cls.get('adjacent_text', u''),
                'reason': cls.get('reason')
            })

    attach_short_exterior_returns(boundary_items)
    exterior_direction_count = count_exterior_directions_after_merge(boundary_items)
    corner_add = CORNER_ADD_TEMP if exterior_direction_count >= CORNER_MIN_DIRECTIONS else 0.0
    t_in = base_t + corner_add

    grouped = {}
    group_segments = defaultdict(list)
    rows = []
    wall_row_by_group_key = {}

    for item in boundary_items:
        wall = item.get('wall')
        kind = item.get('kind')
        adj = item.get('adjacent')
        curve = item.get('curve')
        length_m = item.get('length_m') or 0.0
        if length_m <= 0.01:
            continue

        if kind == u'НС':
            ogr = u'НС'
            adjacent_text = u'наружный воздух'
            dt = t_in - OUTDOOR_TEMP
            count_this = True
        else:
            if adj is not None and is_common_room(adj):
                ogr = u'ВНС'
                adjacent_text = room_label(adj)
                dt = t_in - COMMON_TEMP
                count_this = True
            elif item.get('assumed_common'):
                ogr = u'ВНС'
                adjacent_text = item.get('adjacent_text') or UNKNOWN_INTERNAL_ADJACENT_TEXT
                dt = t_in - COMMON_TEMP
                count_this = True
            else:
                count_this = False
                if adj is not None:
                    skipped_internal.append({
                        'element_id': elem_id(wall),
                        'adjacent': room_label(adj),
                        'reason': u'смежное помещение не относится к МОП/коридорам/общим помещениям',
                        'wall': wall_type_name(wall)
                    })
        if not count_this:
            continue

        merge_target = item.get('merge_to_item') if ogr == u'НС' else None
        key_curve = merge_target.get('curve') if merge_target is not None else curve
        key_wall = merge_target.get('wall') if merge_target is not None else wall
        key_orient = merge_target.get('orient') if merge_target is not None else (item.get('orient') or u'')

        if ogr == u'НС':
            key = line_group_key_no_wall(key_curve, ogr, adjacent_text, dt, K_WALL)
        else:
            key = line_group_key(key_curve, ogr, adjacent_text, dt, K_WALL, key_wall)

        item['group_key'] = key
        if merge_target is not None:
            try:
                merge_target['group_key'] = key
            except:
                pass

        if key not in grouped:
            grouped[key] = {
                'ogr': ogr,
                'source': u'Стена',
                'element_ids': set(),
                'adjacent': adjacent_text,
                'a_m': 0.0,
                'b_m': height_m,
                'k': K_WALL,
                'dt': dt,
                'orient': key_orient,
                'wall_type': wall_type_name(key_wall),
                'segments': 0,
                'axis_infos': [],
                'extra_length_m': 0.0
            }
        grouped[key]['a_m'] += length_m
        if merge_target is not None:
            grouped[key]['extra_length_m'] += length_m
        else:
            axis_info = wall_axis_interval_info(curve, wall)
            if axis_info is not None:
                grouped[key]['axis_infos'].append(axis_info)
        grouped[key]['element_ids'].add(elem_id(wall))
        grouped[key]['segments'] += 1
        group_segments[key].append(item)
        used_wall_seg_count += 1

    # Стены.
    for key in grouped:
        g = grouped[key]
        ids = sorted(list(g['element_ids']))
        eid_text = u', '.join([ustr(x) for x in ids])
        reasons = []
        try:
            for it in group_segments.get(key, []):
                rs = it.get('reason') or u''
                if rs and rs not in reasons:
                    reasons.append(rs)
        except:
            pass
        reason_text = u'; '.join(reasons[:2])
        note = u'Сгруппированная линия границы помещения: {0} сегм.; стены ElementId {1}. Базовая стена: {2}'.format(g['segments'], eid_text, g['wall_type'])
        if reason_text:
            note += u'; ' + reason_text
        extra_m = safe_float(g.get('extra_length_m'), 0.0)
        base_fallback_m = max(0.0, g['a_m'] - extra_m)
        a_calc_m = centerline_length_from_axis_infos(g.get('axis_infos', []), base_fallback_m, g['ogr'] == u'НС') + extra_m
        t_adj_row = OUTDOOR_TEMP if g['ogr'] == u'НС' else COMMON_TEMP
        row = make_calc_row(
            room, g['ogr'], g['source'], eid_text, g['adjacent'],
            a_calc_m, g['b_m'], g['k'], g['dt'], g['orient'], note,
            u'граница помещения из связанной АР-модели', 0.0, ORIENT_ADD, t_adj_row
        )
        row['_group_key'] = key
        wall_row_by_group_key[key] = row
        rows.append(row)

    # Окна/витражи/двери.
    used_opening_ids = set()
    openings = collect_openings_for_room(room, link_doc, boundary_items, used_opening_ids)
    for op_info in openings:
        op = op_info.get('element')
        host_item = op_info.get('boundary_item')
        if op is None or host_item is None:
            continue

        a, b, size_src = opening_size_m(op)
        if a <= 0.01 or b <= 0.01:
            continue

        host_kind = host_item.get('kind')
        host_adj = host_item.get('adjacent')

        if is_door(op):
            # Дверь считаем только если она действительно ведёт в МОП/коридор/лестницу.
            # Монтажные проёмы и двери в фасадной стене по умолчанию не считаем как дверь на -33.
            if is_false_door_candidate(op):
                continue
            if host_kind == u'ВНС' and host_adj is not None and is_common_room(host_adj):
                ogr = u'Д'
                k = K_DOOR
                inf = 0.0
                adjacent_text = room_label(host_adj)
                t_adj_row = COMMON_TEMP
                dt = t_in - t_adj_row
            else:
                continue
        else:
            # Окна/витражи считаем только на наружной границе помещения.
            # Это защищает от перескока окна/витража в соседнее помещение по HostId стены.
            if host_kind != u'НС':
                continue
            ogr = u'ОК'
            k = K_WINDOW
            inf = WINDOW_INFILTRATION
            adjacent_text = u'наружный воздух'
            t_adj_row = OUTDOOR_TEMP
            dt = t_in - t_adj_row

        note = elem_name(op) + u'; ' + size_src
        row = make_calc_row(
            room, ogr, u'Проём', ustr(elem_id(op)), adjacent_text,
            a, b, k, dt, u'', note,
            u'окно/витраж/дверь из связанной АР-модели', inf, ORIENT_ADD, t_adj_row
        )
        rows.append(row)
        subtract_opening_area_from_host_wall(host_item, a * b, wall_row_by_group_key)

    # КИВ в рабочей модели.
    kivs = collect_current_model_kivs_for_room(room, link_inst)
    for kiv in kivs:
        row = {
            'room_id': eid_int(room.Id),
            'room_number': room_number(room),
            'room_name': room_name(room),
            'room_label': room_label(room),
            'ogr': u'КИВ',
            'source': u'КИВ',
            'element_id': ustr(elem_id(kiv)),
            'adjacent': u'',
            't_adj': None,
            'a_m': 0.0,
            'b_m': 0.0,
            'f_m2': 0.0,
            'k': 0.0,
            'dt': 0.0,
            'q_w': KIV_W,
            'orient_add': 0.0,
            'inf': 0.0,
            'dob': 1.0,
            'q_total': KIV_W,
            'orient': u'',
            'note': elem_name(kiv),
            'formula': u'Qобщ={0} Вт на 1 КИВ'.format(fmt(KIV_W, 1)),
            'formula_source': u'КИВ найден в текущей рабочей модели'
        }
        rows.append(row)

    rows.sort(key=lambda r: (r['room_number'], r['ogr'], r['source'], ustr(r['element_id'])))
    return {
        'room': room,
        'rows': rows,
        'skipped_internal': skipped_internal,
        'raw_wall_count': raw_wall_count,
        'calc_wall_groups': len(grouped),
        'exterior_direction_count': exterior_direction_count,
        'corner_add': corner_add,
        't_in_base': base_t,
        't_in': t_in,
        't_source': t_source,
        'kiv_count': len(kivs),
        'openings_count': len(openings),
        'height_m': height_m
    }



def excel_color(r, g, b):
    return int(r) + int(g) * 256 + int(b) * 65536


# Константы Excel. Числа используем напрямую, чтобы не зависеть от Excel Interop enum.
XL_CENTER = -4108
XL_LEFT = -4131
XL_RIGHT = -4152
XL_CONTINUOUS = 1
XL_THIN = 2
XL_MEDIUM = -4138
XL_LANDSCAPE = 2
XL_OPENXML_WORKBOOK = 51


def get_excel_app():
    # Надежный запуск Excel через COM. Не используем Excel.ApplicationClass,
    # потому что в pyRevit/IronPython на части машин этот класс недоступен.
    import System
    from System import Type, Activator
    excel_type = Type.GetTypeFromProgID('Excel.Application')
    if excel_type is None:
        raise Exception(u'Microsoft Excel не найден через COM ProgID Excel.Application')
    return Activator.CreateInstance(excel_type)


def xl_range(ws, r1, c1, r2=None, c2=None):
    if r2 is None:
        r2 = r1
    if c2 is None:
        c2 = c1
    return ws.Range[ws.Cells[r1, c1], ws.Cells[r2, c2]]


def xl_set_value(ws, r, c, value):
    try:
        ws.Cells[r, c].Value2 = value
    except:
        ws.Cells[r, c].Value = value


def xl_write_row(ws, row, values):
    for i, v in enumerate(values):
        xl_set_value(ws, row, i + 1, v)


def xl_border(rng, weight=XL_THIN):
    try:
        rng.Borders.LineStyle = XL_CONTINUOUS
        rng.Borders.Weight = weight
    except:
        pass


def xl_style_range(rng, fill=None, bold=False, h_align=XL_CENTER, v_align=XL_CENTER,
                   font_size=10, wrap=False, border_weight=XL_THIN):
    try:
        rng.Font.Name = 'Arial'
        rng.Font.Size = font_size
        rng.Font.Bold = bool(bold)
        rng.HorizontalAlignment = h_align
        rng.VerticalAlignment = v_align
        rng.WrapText = bool(wrap)
        if fill is not None:
            rng.Interior.Color = fill
        xl_border(rng, border_weight)
    except:
        pass


def xl_set_row_height(ws, r, h=15):
    try:
        ws.Rows[r].RowHeight = h
    except:
        pass


def xl_num_format(rng, fmt_text):
    try:
        rng.NumberFormat = fmt_text
    except:
        pass


def calc_apartment_total(apt):
    total = 0.0
    for rr in apt.get('room_results', []):
        for row in rr.get('rows', []):
            total += safe_float(row.get('q_total'), 0.0)
    return total


def apartment_short_name(name):
    s = ustr(name).strip()
    s = re.sub(ur'^\s*квартира\s+', u'', s, flags=re.I)
    return s


def report_title_from_link(link_name):
    s = ustr(link_name)
    s = re.sub(ur'\.rvt$', u'', s, flags=re.I).strip()
    if not s:
        s = u'Теплопотери'
    return s


def row_temperature_for_adjacent(data):
    if data is not None and data.get('t_adj') is not None:
        try:
            return round(float(data.get('t_adj')), 0)
        except:
            pass
    ogr = data.get('ogr', u'')
    if ogr in [u'НС', u'ОК']:
        return round(float(OUTDOOR_TEMP), 0)
    if ogr in [u'ВНС', u'Д']:
        return round(float(COMMON_TEMP), 0)
    return u''


def make_excel_line(room, rr, data, apartment_name, first_room, first_row_in_room):
    return [
        apartment_name if (first_room and first_row_in_room) else u'',
        room_name(room) if first_row_in_room else u'',
        round(float(rr['t_in']), 0) if first_row_in_room else u'',
        row_temperature_for_adjacent(data),
        data.get('ogr', u''),
        data.get('orient', u''),
        round(float(data.get('a_m', 0.0)), 1) if data.get('ogr') != u'КИВ' else u'',
        round(float(data.get('b_m', 0.0)), 1) if data.get('ogr') != u'КИВ' else u'',
        round(float(data.get('f_m2', 0.0)), 1) if data.get('ogr') != u'КИВ' else u'',
        round(float(data.get('k', 0.0)), 2) if data.get('ogr') != u'КИВ' else u'',
        round(float(data.get('dt', 0.0)), 0) if data.get('ogr') != u'КИВ' else u'',
        round(float(data.get('q_w', 0.0)), 0),
        round(float(data.get('orient_add', 0.0)), 1) if data.get('ogr') != u'КИВ' else u'',
        round(float(data.get('inf', 0.0)), 1) if data.get('ogr') != u'КИВ' else u'',
        round(float(data.get('dob', 0.0)), 1) if data.get('ogr') != u'КИВ' else u'',
        round(float(data.get('q_total', 0.0)), 0)
    ]


def apply_body_number_formats(ws, r1, r2):
    if r2 < r1:
        return
    for c in [3, 4, 11, 12, 16]:
        xl_num_format(xl_range(ws, r1, c, r2, c), '0')
    for c in [7, 8, 9, 13, 14, 15]:
        xl_num_format(xl_range(ws, r1, c, r2, c), '0.0')
    xl_num_format(xl_range(ws, r1, 10, r2, 10), '0.00')


def save_xlsx_report(apartment_results, level_name, link_name, apartment_mode, height_source):
    # Настоящий XLSX без шаблона, без HTML и без псевдо-.xls.
    # Стиль повторяет присланную таблицу: зеленая строка объекта, белая шапка,
    # голубые названия помещений, итог помещения только в последней зеленой ячейке.
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    fname = u'Теплопотери_OV_{0}_отчет.xlsx'.format(strip_bad_filename_chars(level_name))
    path = os.path.join(desktop, fname)
    if os.path.exists(path):
        try:
            os.remove(path)
        except:
            # Если файл открыт, сохраняем рядом с индексом.
            base, ext = os.path.splitext(path)
            i = 1
            while os.path.exists(base + u'_{0}'.format(i) + ext):
                i += 1
            path = base + u'_{0}'.format(i) + ext

    headers = [
        u'№п/п', u'Наим.пом.', u'tв', u'tн', u'Огр.', u'Ориент.',
        u'a,м', u'в,м', u'F,м²', u'к', u'Δt', u'Q,Вт', u'Ор.доб.', u'Инф.', u'Доб.', u'Qобщ.'
    ]

    GREEN = excel_color(146, 208, 80)       # #92D050
    BLUE = excel_color(184, 204, 228)       # #B8CCE4
    WHITE = excel_color(255, 255, 255)

    xl = None
    wb = None
    try:
        xl = get_excel_app()
        xl.Visible = False
        xl.DisplayAlerts = False

        wb = xl.Workbooks.Add()
        ws = wb.Worksheets[1]
        try:
            ws.Name = u'Расчет'
        except:
            pass

        # Удаляем лишние листы.
        try:
            while wb.Worksheets.Count > 1:
                wb.Worksheets[wb.Worksheets.Count].Delete()
        except:
            pass

        # Компактные ширины как на примере.
        widths = [7, 11, 6, 7, 6, 8, 7, 7, 8, 6, 6, 8, 8, 6, 6, 8]
        for i, w in enumerate(widths):
            try:
                ws.Columns[i + 1].ColumnWidth = w
            except:
                pass

        try:
            ws.Cells.Font.Name = 'Arial'
            ws.Cells.Font.Size = 10
            ws.Rows.RowHeight = 15
        except:
            pass

        row = 1

        # Верхняя зеленая строка объекта.
        title_rng = xl_range(ws, row, 1, row, 16)
        title_rng.Merge()
        xl_set_value(ws, row, 1, report_title_from_link(link_name))
        xl_style_range(title_rng, fill=GREEN, bold=True, h_align=XL_CENTER)
        xl_set_row_height(ws, row, 15)
        row += 1

        # Белая шапка.
        xl_write_row(ws, row, headers)
        xl_style_range(xl_range(ws, row, 1, row, 16), fill=WHITE, bold=False, h_align=XL_CENTER)
        xl_set_row_height(ws, row, 15)
        row += 1

        grand_total = 0.0

        for apt in apartment_results:
            apt_name = apartment_short_name(apt.get('name', u''))
            room_results = apt.get('room_results', [])
            grand_total += calc_apartment_total(apt)

            first_room_in_apartment = True

            for rr in room_results:
                room = rr['room']
                rows = list(rr.get('rows', []))
                room_total = sum([safe_float(r.get('q_total'), 0.0) for r in rows])
                room_start_row = row

                if not rows:
                    line = [
                        apt_name if first_room_in_apartment else u'',
                        room_name(room),
                        round(float(rr['t_in']), 0),
                        u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u'',
                        round(float(room_total), 0)
                    ]
                    xl_write_row(ws, row, line)
                    xl_style_range(xl_range(ws, row, 1, row, 16), fill=WHITE, h_align=XL_CENTER)
                    xl_range(ws, row, 2).Interior.Color = BLUE
                    xl_set_row_height(ws, row, 15)
                    row += 1
                    first_room_in_apartment = False
                else:
                    for idx, data in enumerate(rows):
                        line = make_excel_line(room, rr, data, apt_name, first_room_in_apartment, idx == 0)
                        xl_write_row(ws, row, line)
                        xl_style_range(xl_range(ws, row, 1, row, 16), fill=WHITE, h_align=XL_CENTER)
                        xl_set_row_height(ws, row, 15)
                        row += 1

                    first_room_in_apartment = False

                    # Объединение №п/п, Наим.пом. и tв на строки помещения.
                    room_end_row = row - 1
                    if room_end_row >= room_start_row:
                        for c in [1, 2, 3]:
                            rng = xl_range(ws, room_start_row, c, room_end_row, c)
                            try:
                                if room_end_row > room_start_row:
                                    rng.Merge()
                            except:
                                pass
                            xl_style_range(rng, fill=BLUE if c == 2 else WHITE, h_align=XL_CENTER)

                # Голубая заливка только в колонке имени помещения.
                try:
                    xl_range(ws, room_start_row, 2, row - 1, 2).Interior.Color = BLUE
                except:
                    pass

                apply_body_number_formats(ws, room_start_row, row - 1)

                # Итог помещения: строка белая, зеленая только последняя ячейка Qобщ.
                xl_style_range(xl_range(ws, row, 1, row, 15), fill=WHITE, h_align=XL_CENTER)
                xl_set_value(ws, row, 16, round(float(room_total), 0))
                total_cell = xl_range(ws, row, 16)
                xl_style_range(total_cell, fill=GREEN, bold=True, h_align=XL_RIGHT)
                xl_num_format(total_cell, '0')
                xl_set_row_height(ws, row, 15)
                row += 1

        # Общий итог внизу. Сделан зеленым, не меняет структуру комнат.
        total_rng = xl_range(ws, row, 1, row, 15)
        total_rng.Merge()
        xl_set_value(ws, row, 1, u'ИТОГО ПО ЭТАЖУ')
        xl_set_value(ws, row, 16, round(float(grand_total), 0))
        xl_style_range(xl_range(ws, row, 1, row, 16), fill=GREEN, bold=True, h_align=XL_RIGHT)
        xl_num_format(xl_range(ws, row, 16), '0')
        xl_set_row_height(ws, row, 15)

        # Параметры страницы и закрепление шапки.
        try:
            ws.Application.ActiveWindow.SplitRow = 2
            ws.Application.ActiveWindow.FreezePanes = True
        except:
            pass
        try:
            ws.PageSetup.Orientation = XL_LANDSCAPE
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False
        except:
            pass

        try:
            ws.Range['A1'].Select()
        except:
            pass

        wb.SaveAs(path, XL_OPENXML_WORKBOOK)
        wb.Close(False)
        xl.Quit()
        return path, grand_total
    except:
        try:
            if wb is not None:
                wb.Close(False)
        except:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except:
            pass
        raise

def print_window_report(apartment_results, total, path, level_name, apartment_mode, height_source):
    output.print_md(u'# Теплопотери ОВ — расчет квартир этажа')
    output.print_md(u'**Этаж:** {0}'.format(level_name))
    output.print_md(u'**Определение квартир:** {0}'.format(apartment_mode))
    output.print_md(u'**Высота стен:** {0}'.format(height_source))
    output.print_md(u'**Итого по этажу:** **{0} Вт**'.format(fmt0(total)))
    output.print_md(u'**Excel-отчёт:** `{0}`'.format(path))

    output.print_md(u'## Сводка по квартирам')
    output.print_md(u'| Квартира | Помещений | Qобщ., Вт |')
    output.print_md(u'|---|---:|---:|')
    for apt in apartment_results:
        apt_total = sum([sum([r['q_total'] for r in rr['rows']]) for rr in apt['room_results']])
        output.print_md(u'| {0} | {1} | {2} |'.format(apt['name'], len(apt['room_results']), fmt0(apt_total)))

    output.print_md(u'## Что принято в расчете')
    output.print_md(u'- Наружные стены и окна считаются с tн = {0} °C.'.format(fmt(OUTDOOR_TEMP, 0)))
    output.print_md(u'- Внутренние стены к найденным МОП/коридорам/лестничным клеткам считаются с t = {0} °C.'.format(fmt(COMMON_TEMP, 0)))
    output.print_md(u'- Если соседнее помещение не найдено, но стена распознана как внутренняя, она также считается как ВНС с t = {0} °C, а не как наружная на -33 °C.'.format(fmt(COMMON_TEMP, 0)))
    output.print_md(u'- Внутренние стены к квартирным помещениям пропускаются.')
    output.print_md(u'- Если у помещения 2 и более наружных направления, к tв добавляется +{0} °C.'.format(fmt(CORNER_ADD_TEMP, 0)))
    output.print_md(u'- КИВ ищется в текущей рабочей модели, помещения и ограждения — в связанной АР-модели.')


def open_file(path):
    try:
        os.startfile(path)
    except:
        pass


def main():
    output.close_others()
    link_inst, link_doc, level, levels = choose_link_and_level()
    rooms_on_level = get_rooms_on_level(link_doc, level)
    if not rooms_on_level:
        forms.alert(u'На выбранном уровне в связанной модели не найдены помещения.', exitscript=True)

    height_m, height_source = next_level_height_m(level, levels)
    height_source_full = u'{0}, h={1} м'.format(height_source, fmt(height_m, 3))

    shared_boundary_map = build_shared_boundary_map(rooms_on_level, link_doc)
    apartments, common_rooms, apartment_mode = determine_apartments(rooms_on_level, link_doc, shared_boundary_map)
    if not apartments:
        forms.alert(u'Не удалось определить квартиры на выбранном этаже. Проверь имена помещений или параметры квартиры.', exitscript=True)

    apartment_results = []
    processed = 0
    with forms.ProgressBar(title=u'Теплопотери ОВ: расчет квартир', cancellable=True) as pb:
        total_rooms = sum([len(a['rooms']) for a in apartments])
        for apt in apartments:
            room_results = []
            for r in apt['rooms']:
                if pb.cancelled:
                    script.exit()
                processed += 1
                pb.update_progress(processed, total_rooms)
                rr = calculate_room(r, link_doc, link_inst, rooms_on_level, height_m, shared_boundary_map)
                room_results.append(rr)
            apartment_results.append({'name': apt['name'], 'room_results': room_results, 'source': apt.get('source', u'')})

    report_path, total = save_xlsx_report(apartment_results, level.Name, link_doc.Title, apartment_mode, height_source_full)
    print_window_report(apartment_results, total, report_path, level.Name, apartment_mode, height_source_full)
    open_file(report_path)

    forms.alert(
        u'Расчет теплопотерь по квартирам выполнен.\n\n'
        u'Этаж: {0}\n'
        u'Квартир: {1}\n'
        u'Итого: {2} Вт\n\n'
        u'Компактный XLSX-отчет открыт и сохранен на рабочем столе.'.format(level.Name, len(apartment_results), fmt0(total)),
        title=u'Теплопотери ОВ'
    )


if __name__ == '__main__':
    try:
        main()
    except OperationCanceledException:
        pass
    except Exception as ex:
        output.print_md(u'## Ошибка')
        output.print_md(u'```')
        output.print_md(ustr(traceback.format_exc()))
        output.print_md(u'```')
        forms.alert(u'Расчет прерван ошибкой. Подробности выведены в окно pyRevit.', title=u'Теплопотери ОВ')
