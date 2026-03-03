# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms
import csv, os, math
from collections import defaultdict

doc = revit.doc

# =========================================================
# 表 8：各地區日射修正係數 fk（住宅類）
# 方向命名：S/SSW/SW/WSW/W/WNW/NW/NNW/N/NNE/NE/ENE/E/ESE/SE/SSE
# =========================================================
FK_TABLE = {
    "臺北市": {
        "H": 2.31,
        "S": 1.00, "SSW": 1.04, "SW": 1.06, "WSW": 1.05,
        "W": 1.00, "WNW": 0.93, "NW": 0.84, "NNW": 0.76,
        "N": 0.71, "NNE": 0.73, "NE": 0.78, "ENE": 0.84,
        "E": 0.90, "ESE": 0.94, "SE": 0.97, "SSE": 0.99
    },
    "臺中市": {
        "H": 3.29,
        "S": 1.51, "SSW": 1.63, "SW": 1.69, "WSW": 1.68,
        "W": 1.57, "WNW": 1.40, "NW": 1.20, "NNW": 1.03,
        "N": 0.91, "NNE": 0.91, "NE": 0.97, "ENE": 1.06,
        "E": 1.15, "ESE": 1.25, "SE": 1.34, "SSE": 1.43
    },
    "花蓮市": {
        "H": 2.86,
        "S": 1.02, "SSW": 1.10, "SW": 1.17, "WSW": 1.20,
        "W": 1.17, "WNW": 1.10, "NW": 0.98, "NNW": 0.86,
        "N": 0.75, "NNE": 0.77, "NE": 0.84, "ENE": 0.92,
        "E": 0.98, "ESE": 1.03, "SE": 1.04, "SSE": 1.03
    },
    "高雄市": {
        "H": 3.75,
        "S": 1.70, "SSW": 1.83, "SW": 1.90, "WSW": 1.88,
        "W": 1.76, "WNW": 1.58, "NW": 1.36, "NNW": 1.18,
        "N": 1.06, "NNE": 1.06, "NE": 1.12, "ENE": 1.20,
        "E": 1.29, "ESE": 1.39, "SE": 1.48, "SSE": 1.58
    },
    "臺東市": {
        "H": 4.03,
        "S": 1.63, "SSW": 1.74, "SW": 1.82, "WSW": 1.81,
        "W": 1.71, "WNW": 1.53, "NW": 1.30, "NNW": 1.08,
        "N": 0.94, "NNE": 0.95, "NE": 1.05, "ENE": 1.18,
        "E": 1.32, "ESE": 1.43, "SE": 1.51, "SSE": 1.57
    }
}

DIR16 = ["N","NNE","NE","ENE","E","ESE","SE","SSE","S","SSW","SW","WSW","W","WNW","NW","NNW"]

# =========================================================
# 工具函式
# =========================================================
def vec_to_azimuth_deg(v):
    """XY 平面向量 -> 方位角（度），0=北、90=東"""
    ang = math.degrees(math.atan2(v.X, v.Y))
    return ang + 360.0 if ang < 0 else ang

def az_to_dir16(az):
    idx = int((az + 11.25) // 22.5) % 16
    return DIR16[idx]

def transform_vector(t, v):
    bx, by, bz = t.BasisX, t.BasisY, t.BasisZ
    return DB.XYZ(
        bx.X*v.X + by.X*v.Y + bz.X*v.Z,
        bx.Y*v.X + by.Y*v.Y + bz.Y*v.Z,
        bx.Z*v.X + by.Z*v.Y + bz.Z*v.Z
    )

def get_project_position_angle(doc):
    """Project North -> True North 的角度（弧度，帶正負號）"""
    pl = doc.ActiveProjectLocation
    pp = pl.GetProjectPosition(DB.XYZ(0,0,0))
    return pp.Angle  # radians, sign included

def to_m(x_internal):
    return DB.UnitUtils.ConvertFromInternalUnits(x_internal, DB.UnitTypeId.Meters)

def to_m2(a_internal):
    return DB.UnitUtils.ConvertFromInternalUnits(a_internal, DB.UnitTypeId.SquareMeters)

def get_level_name_from_wall(wall):
    try:
        p = wall.get_Parameter(DB.BuiltInParameter.WALL_BASE_CONSTRAINT)
        if p:
            lvl = doc.GetElement(p.AsElementId())
            return lvl.Name if lvl else ""
    except:
        pass
    return ""

def get_level_name_from_elem(elem):
    try:
        lvl = doc.GetElement(elem.LevelId)
        return lvl.Name if lvl else ""
    except:
        return ""

def get_type_mark(elem_type):
    p = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
    if p:
        s = p.AsString()
        if s:
            return s
    return None

def get_window_width_height_m(win):
    sym = win.Symbol
    # 通用抓法
    pw = sym.get_Parameter(DB.BuiltInParameter.WINDOW_WIDTH)
    ph = sym.get_Parameter(DB.BuiltInParameter.WINDOW_HEIGHT)
    if pw and ph:
        return to_m(pw.AsDouble()), to_m(ph.AsDouble())
    # fallback：找常見參數名
    for wn in ["Width","寬度","寬"]:
        p = sym.LookupParameter(wn)
        if p and p.StorageType == DB.StorageType.Double:
            w = to_m(p.AsDouble()); break
    else:
        w = None
    for hn in ["Height","高度","高"]:
        p = sym.LookupParameter(hn)
        if p and p.StorageType == DB.StorageType.Double:
            h = to_m(p.AsDouble()); break
    else:
        h = None
    return w, h

def get_window_facing_vector(win):
    host = win.Host
    if isinstance(host, DB.Wall):
        v = host.Orientation
        v = DB.XYZ(v.X, v.Y, 0)
        if v.GetLength() > 1e-6:
            return v.Normalize()
    try:
        v = win.FacingOrientation
        v = DB.XYZ(v.X, v.Y, 0)
        if v.GetLength() > 1e-6:
            return v.Normalize()
    except:
        pass
    return None

def get_wall_facing_vector(wall):
    try:
        v = wall.Orientation
        v = DB.XYZ(v.X, v.Y, 0)
        if v.GetLength() > 1e-6:
            return v.Normalize()
    except:
        pass
    return None

def get_curtainwall_width_height_m(wall):
    """DW 直接用：牆長(CURVE_ELEM_LENGTH) × 牆高(WALL_USER_HEIGHT_PARAM)"""
    p_len = wall.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
    p_hgt = wall.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)  # Unconnected Height
    if not p_len or not p_hgt:
        return None, None
    return to_m(p_len.AsDouble()), to_m(p_hgt.AsDouble())

# =========================================================
# CSV（IronPython 相容）：binary + UTF-8 BOM
# =========================================================
def _to_utf8(x):
    try:
        if x is None:
            return ""
        if isinstance(x, (int, float)):
            return str(x)
        if isinstance(x, bytes):
            return x
        return unicode(x).encode("utf-8")  # noqa
    except:
        try:
            return str(x)
        except:
            return ""

def write_csv_utf8_bom(path, headers, rows):
    bom = u"\ufeff".encode("utf-8")
    with open(path, "wb") as f:
        f.write(bom)
        wr = csv.writer(f)
        wr.writerow([_to_utf8(h) for h in headers])
        for r in rows:
            wr.writerow([_to_utf8(v) for v in r])

def safe_name(x, fallback=""):
    try:
        if x is None:
            return fallback
        if isinstance(x, DB.ElementId):
            el = doc.GetElement(x)
            return el.Name if el else fallback
        return x.Name
    except:
        return fallback


# =========================================================
# UI：地區 / 真北 / 計算對象 / 輸出資料夾
# =========================================================
city = forms.SelectFromList.show(
    sorted(FK_TABLE.keys()),
    title="選擇氣候地區（表 8 fk）",
    multiselect=False
)
if not city:
    forms.alert("已取消。", exitscript=True)

use_true_north = forms.alert(
    "要用「真北」計算方位嗎？\nYes=真北 / No=專案北",
    yes=True, no=True
)

target = forms.CommandSwitchWindow.show(
    ["Windows", "CurtainWall(DW)", "Windows+CurtainWall(DW)"],
    message="選擇要計算的對象："
)
if not target:
    forms.alert("已取消。", exitscript=True)

out_dir = forms.pick_folder(title="選擇輸出資料夾")
if not out_dir:
    forms.alert("已取消。", exitscript=True)

if not os.path.exists(out_dir):
    os.makedirs(out_dir)

out_rows = os.path.join(out_dir, "Req_D1_rows.csv")
out_inst = os.path.join(out_dir, "Req_instances.csv")

ang = get_project_position_angle(doc) if use_true_north else None
#（可用來檢核）
# forms.alert("ProjectPosition.Angle(deg) = {}".format(round(math.degrees(ang), 3)) if ang else "0")

# =========================================================
# 表頭（summary + instances，跟你窗的一樣）
# =========================================================
inst_headers = [
    "ElementId","Level","窗編號(TypeMark)",
    "Width_m","Height_m","Agi_m2",
    "AzimuthDeg(0=N)","方位16","fk",
    "Ki(固定)","Agi_fk_Ki"
]

d1_headers = [
    "方位","日射修正係數fk","窗編號",
    "寬(m)","高(m)","每種窗Agi",
    "數量ni","窗面積ΣAgi","外遮陽Ki(固定)",
    "外殼等價開窗面積ΣAgi*fk*Ki(m2)"
]

KI_FIXED = 1.0

# =========================================================
# 收集元素
# =========================================================
items = []

if target in ["Windows", "Windows+CurtainWall(DW)"]:
    wins = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_Windows) \
        .WhereElementIsNotElementType() \
        .ToElements()
    for w in wins:
        items.append(("WINDOW", w))

if target in ["CurtainWall(DW)", "Windows+CurtainWall(DW)"]:
    walls = DB.FilteredElementCollector(doc) \
        .OfClass(DB.Wall) \
        .WhereElementIsNotElementType() \
        .ToElements()
    for wall in walls:
        try:
            wt = doc.GetElement(wall.GetTypeId())
            if isinstance(wt, DB.WallType) and wt.Kind == DB.WallKind.Curtain:
                items.append(("CWALL", wall))
        except:
            pass

if not items:
    forms.alert("找不到可計算的元素。", exitscript=True)

# =========================================================
# 計算 + 彙總
# 需求：樓層不影響 summary，只用 方位16 + fk + 窗編號 + 寬高
# =========================================================
groups = defaultdict(lambda: {"count":0, "agi_sum":0.0, "agi_fk_ki_sum":0.0})
inst_rows = []
skipped = 0

for kind, e in items:
    if kind == "WINDOW":
        facing = get_window_facing_vector(e)
        if not facing:
            skipped += 1
            continue

        if use_true_north:
            facing = transform_vector(doc.ActiveProjectLocation.GetTotalTransform(), facing)  # 保留你的總變換
            # 但你真北換算以 Angle 為準（已驗證用 ang）
            facing = DB.XYZ(facing.X, facing.Y, 0).Normalize()
            facing = DB.XYZ(
                facing.X*math.cos(ang) - facing.Y*math.sin(ang),
                facing.X*math.sin(ang) + facing.Y*math.cos(ang),
                0
            ).Normalize()

        az = vec_to_azimuth_deg(facing)
        dir16 = az_to_dir16(az)
        fk = FK_TABLE[city].get(dir16, None)
        if fk is None:
            skipped += 1
            continue

        sym = e.Symbol
        type_mark = get_type_mark(sym) or sym.Name

        level_name = get_level_name_from_elem(e)

        w_m, h_m = get_window_width_height_m(e)
        if w_m is None or h_m is None:
            skipped += 1
            continue

    else:  # CWALL (DW)
        facing = get_wall_facing_vector(e)
        if not facing:
            skipped += 1
            continue

        if use_true_north:
            # 直接用 Angle 旋轉（用 ang，不手動反號）
            facing = DB.XYZ(
                facing.X*math.cos(ang) - facing.Y*math.sin(ang),
                facing.X*math.sin(ang) + facing.Y*math.cos(ang),
                0
            ).Normalize()

        az = vec_to_azimuth_deg(facing)
        dir16 = az_to_dir16(az)
        fk = FK_TABLE[city].get(dir16, None)
        if fk is None:
            skipped += 1
            continue

        wt = doc.GetElement(e.GetTypeId())
        type_mark = get_type_mark(wt) or safe_name(wt, "DW_TYPE_MISSING")

        level_name = get_level_name_from_wall(e)

        w_m, h_m = get_curtainwall_width_height_m(e)
        if w_m is None or h_m is None:
            skipped += 1
            continue

    agi = w_m * h_m
    agi_fk_ki = agi * fk * KI_FIXED

    inst_rows.append([
        e.Id.IntegerValue, level_name, type_mark,
        round(w_m, 3), round(h_m, 3), round(agi, 3),
        round(az, 2), dir16, round(fk, 2),
        1, round(agi_fk_ki, 3)
    ])

    gkey = (dir16, round(fk, 2), type_mark, round(w_m, 3), round(h_m, 3))
    groups[gkey]["count"] += 1
    groups[gkey]["agi_sum"] += agi
    groups[gkey]["agi_fk_ki_sum"] += agi_fk_ki

# =========================================================
# 組 summary rows
# =========================================================
d1_rows = []
for (dir16, fk2, type_mark, w_m, h_m), data in sorted(groups.items()):
    ni = data["count"]
    agi_each = w_m * h_m
    agi_sum = data["agi_sum"]  # ΣAgi
    d1_rows.append([
        dir16, fk2, type_mark,
        w_m, h_m, round(agi_each, 3),
        ni, round(agi_sum, 3), 1,
        round(data["agi_fk_ki_sum"], 3)
    ])

# =========================================================
# 寫檔
# =========================================================
write_csv_utf8_bom(out_rows, d1_headers, d1_rows)
write_csv_utf8_bom(out_inst, inst_headers, inst_rows)

forms.alert(
    "完成輸出：\n"
    "- Summary：{}\n"
    "- Instances：{}\n\n"
    "統計：\n"
    "- 計算元素數：{}\n"
    "- 略過：{}\n\n"
    "設定：{} / 方位基準：{}".format(
        out_rows, out_inst,
        len(items), skipped,
        city, ("真北" if use_true_north else "專案北")
    )
)
