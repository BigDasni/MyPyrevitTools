# -*- coding: utf-8 -*-
"""
ChamferLines — Extend selected line segments to their mutual intersection point
to close open corners.

Supported line types:
  - Detail Lines          (OST_Lines)
  - Room Separation Lines (OST_RoomSeparationLines)
  - Area Boundary Lines   (OST_AreaSchemeLines)

Workflow:
  1. Pre-select lines in Revit, OR run and pick interactively.
  2. Tool asks for a gap-detection threshold (default 100 mm).
  3. Near-endpoint pairs are found across different lines.
  4. For each pair, the mathematical intersection of the two lines is computed.
  5. Each line's near-endpoint is moved to that intersection (SetGeometryCurve).
     Falls back to delete-and-recreate for boundary line types.
"""

from pyrevit import forms, revit, script
from pyrevit import DB
from Autodesk.Revit.UI.Selection import ObjectType as UIObjectType

doc   = revit.doc
uidoc = revit.uidoc

# ─────────────────────────────────────────────────────────────
# Supported category IDs
# ─────────────────────────────────────────────────────────────
CAT_DETAIL   = int(DB.BuiltInCategory.OST_Lines)
CAT_ROOM_SEP = int(DB.BuiltInCategory.OST_RoomSeparationLines)
CAT_AREA_BND = int(DB.BuiltInCategory.OST_AreaSchemeLines)

SUPPORTED_CAT_IDS = {CAT_DETAIL, CAT_ROOM_SEP, CAT_AREA_BND}

# ─────────────────────────────────────────────────────────────
# Geometry helpers
# ─────────────────────────────────────────────────────────────

def cross2d(v1, v2):
    """Z-component of the 2-D cross product of two XYZ vectors."""
    return v1.X * v2.Y - v1.Y * v2.X


def find_line_intersection_2d(curve_a, curve_b):
    """
    Compute the 2-D (plan) intersection of two Revit Lines extended infinitely.

    Returns an XYZ intersection point, or None if the lines are parallel.
    The Z value is interpolated from curve_a.
    """
    a0 = curve_a.GetEndPoint(0)
    a1 = curve_a.GetEndPoint(1)
    b0 = curve_b.GetEndPoint(0)
    b1 = curve_b.GetEndPoint(1)

    da = DB.XYZ(a1.X - a0.X, a1.Y - a0.Y, 0.0)
    db = DB.XYZ(b1.X - b0.X, b1.Y - b0.Y, 0.0)

    denom = cross2d(da, db)
    if abs(denom) < 1e-9:
        return None  # parallel or coincident

    diff = DB.XYZ(b0.X - a0.X, b0.Y - a0.Y, 0.0)
    t = cross2d(diff, db) / denom

    ix = a0.X + t * da.X
    iy = a0.Y + t * da.Y
    iz = (a0.Z + a1.Z) / 2.0   # keep Z on the line's plane

    return DB.XYZ(ix, iy, iz)


def pts_within(pa, pb, threshold_ft):
    """Return True if the 2-D (plan) distance between pa and pb < threshold."""
    dx = pa.X - pb.X
    dy = pa.Y - pb.Y
    return (dx * dx + dy * dy) < threshold_ft * threshold_ft

# ─────────────────────────────────────────────────────────────
# Element helpers
# ─────────────────────────────────────────────────────────────

def cat_id_of(element):
    return element.Category.Id.IntegerValue


def is_supported_line(element):
    """Return True if the element is a straight-line CurveElement we can handle."""
    try:
        if not isinstance(element, DB.CurveElement):
            return False
        if cat_id_of(element) not in SUPPORTED_CAT_IDS:
            return False
        # Must be a straight line (not arc / spline)
        return isinstance(element.GeometryCurve, DB.Line)
    except Exception:
        return False


def get_owner_view(element):
    """Return the View that owns a view-specific CurveElement, or None."""
    vid = element.OwnerViewId
    if vid == DB.ElementId.InvalidElementId:
        return None
    return doc.GetElement(vid)


def get_or_create_sketch_plane(view):
    """Return the SketchPlane for the view, creating one from GenLevel if needed."""
    sp = view.SketchPlane
    if sp:
        return sp
    if hasattr(view, 'GenLevel') and view.GenLevel:
        sp = DB.SketchPlane.Create(doc, view.GenLevel.Id)
        view.SketchPlane = sp
        return sp
    # Fallback: create from the view's plane
    plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ.Zero)
    return DB.SketchPlane.Create(doc, plane)


def apply_new_geometry(element, new_curve):
    """
    Apply new_curve geometry to an existing CurveElement.

    Strategy:
      1. Try SetGeometryCurve (works for most types, fastest).
      2. If that fails, delete-and-recreate with the appropriate API.

    Returns True on success, False on failure.
    """
    # ── Attempt 1: SetGeometryCurve ──
    try:
        element.SetGeometryCurve(new_curve, False)
        return True
    except Exception:
        pass

    # ── Attempt 2: Delete & Recreate ──
    cat = cat_id_of(element)
    view = get_owner_view(element)
    if view is None:
        return False  # can't recreate without a view context

    try:
        doc.Delete(element.Id)
    except Exception as e:
        print("  [ERROR] Delete failed for element {}: {}".format(element.Id, e))
        return False

    try:
        if cat == CAT_DETAIL:
            doc.Create.NewDetailCurve(view, new_curve)

        elif cat == CAT_ROOM_SEP:
            sp = get_or_create_sketch_plane(view)
            arr = DB.CurveArray()
            arr.Append(new_curve)
            doc.Create.NewRoomBoundaryLines(sp, arr, view)

        elif cat == CAT_AREA_BND:
            sp = get_or_create_sketch_plane(view)
            doc.Create.NewAreaBoundaryLine(sp, new_curve, view)

        return True

    except Exception as e:
        print("  [ERROR] Recreate failed for cat {}: {}".format(cat, e))
        return False

# ─────────────────────────────────────────────────────────────
# Step 1 — Collect selected / picked lines
# ─────────────────────────────────────────────────────────────

sel_ids = uidoc.Selection.GetElementIds()

if not sel_ids or sel_ids.Count == 0:
    # Nothing pre-selected: enter interactive pick mode
    try:
        picked = uidoc.Selection.PickObjects(
            UIObjectType.Element,  # ← 正確 namespace: Autodesk.Revit.UI.Selection
            "請選取需要閉合的線段 (Detail Line / Room Boundary / Area Boundary)，完成後按 Finish"
        )
        sel_ids = [p.ElementId for p in picked]
    except Exception as pick_err:
        # 使用者按 Escape，或發生其他錯誤
        print("[ChamferLines] 選取取消或發生錯誤: {}".format(pick_err))
        script.exit()

# Filter to supported straight-line elements only
all_elements = [doc.GetElement(eid) for eid in sel_ids]
lines = [el for el in all_elements if el is not None and is_supported_line(el)]
skipped_types = len(all_elements) - len(lines)

if skipped_types > 0:
    print("[ChamferLines] {} 個元素不是支援的線型，已略過。".format(skipped_types))

if len(lines) < 2:
    forms.alert(
        "請至少選取 2 條支援的線段。\n"
        "（你選了 {} 個元素，其中 {} 條是支援的直線）\n\n"
        "支援類型：\n"
        "  • Detail Lines (標註線)\n"
        "  • Room Separation Lines (房間邊界)\n"
        "  • Area Boundary Lines (面積邊界)".format(len(all_elements), len(lines)),
        exitscript=True
    )

# ─────────────────────────────────────────────────────────────
# Step 2 — Ask for gap-detection threshold
# ─────────────────────────────────────────────────────────────

threshold_str = forms.ask_for_string(
    prompt="設定端點偵測距離閾值 (mm)：\n\n"
           "兩條線的端點距離小於此值，才會被認定為需要延伸的開口角。\n"
           "(建議值：10 – 200 mm)",
    title="ChamferLines — 閾值設定",
    default="100"
)

if threshold_str is None:
    script.exit()

try:
    threshold_mm = float(threshold_str.strip())
    if threshold_mm <= 0:
        raise ValueError("non-positive")
except Exception:
    forms.alert("無效的數值，改用預設 100 mm。")
    threshold_mm = 100.0

THRESHOLD_FT = threshold_mm / 304.8

# ─────────────────────────────────────────────────────────────
# Step 3 — Build endpoint list & find near-endpoint pairs
# ─────────────────────────────────────────────────────────────
# endpoints: list of (line_index, endpoint_index 0|1, XYZ)

endpoints = []
for i, line in enumerate(lines):
    c = line.GeometryCurve
    endpoints.append((i, 0, c.GetEndPoint(0)))
    endpoints.append((i, 1, c.GetEndPoint(1)))

# Find unique line-pair candidates (one near-endpoint pair per line combination)
# pairs: [(line_idx_A, end_idx_A, line_idx_B, end_idx_B), ...]
pairs = []
seen_line_pairs = set()

n = len(endpoints)
for a in range(n):
    li, ei, pa = endpoints[a]
    for b in range(a + 1, n):
        lj, ej, pb = endpoints[b]
        if li == lj:
            continue                            # same line — skip
        key = (min(li, lj), max(li, lj))
        if key in seen_line_pairs:
            continue                            # already found a pair for this combo
        if pts_within(pa, pb, THRESHOLD_FT):
            pairs.append((li, ei, lj, ej))
            seen_line_pairs.add(key)

if not pairs:
    forms.alert(
        "在 {}mm 閾值內找不到可閉合的端點對。\n\n"
        "請確認：\n"
        "  1. 選取的線段端點彼此距離 < {}mm\n"
        "  2. 線段兩兩相鄰但尚未相交".format(threshold_mm, threshold_mm),
        exitscript=True
    )

# ─────────────────────────────────────────────────────────────
# Step 4 — Calculate new endpoint positions
# ─────────────────────────────────────────────────────────────
# Accumulate per-line endpoint updates.
# A single line may appear in multiple pairs (both corners of an L-shape),
# so we collect all updates first, then apply once.
#
# new_pts[line_idx] = { end_idx: new_XYZ, ... }

new_pts = {}
skipped_parallel = 0

for li, ei, lj, ej in pairs:
    curve_a = lines[li].GeometryCurve
    curve_b = lines[lj].GeometryCurve

    intersection = find_line_intersection_2d(curve_a, curve_b)
    if intersection is None:
        print("  [SKIP] 線段 {} 與 {} 平行，無法求交點。".format(li, lj))
        skipped_parallel += 1
        continue

    # Record the new endpoint for line A (end ei → intersection)
    if li not in new_pts:
        new_pts[li] = {0: curve_a.GetEndPoint(0), 1: curve_a.GetEndPoint(1)}
    new_pts[li][ei] = intersection

    # Record the new endpoint for line B (end ej → intersection)
    if lj not in new_pts:
        new_pts[lj] = {0: curve_b.GetEndPoint(0), 1: curve_b.GetEndPoint(1)}
    new_pts[lj][ej] = intersection

if not new_pts:
    forms.alert("所有端點對均平行，無法延伸。", exitscript=True)

# ─────────────────────────────────────────────────────────────
# Step 5 — Apply inside a single Transaction
# ─────────────────────────────────────────────────────────────

count_ok   = 0
count_fail = 0
count_skip = 0

with revit.Transaction("Chamfer Lines — Extend to Intersection"):
    for line_idx, pt_dict in new_pts.items():
        element = lines[line_idx]
        p0 = pt_dict[0]
        p1 = pt_dict[1]

        # Guard: degenerate line (both endpoints collapsed to the same point)
        if p0.DistanceTo(p1) < 1e-6:
            print("  [SKIP] 線段 {} 長度為 0，跳過。".format(line_idx))
            count_skip += 1
            continue

        try:
            new_line = DB.Line.CreateBound(p0, p1)
        except Exception as e:
            print("  [ERROR] 無法建立新線幾何 {}: {}".format(line_idx, e))
            count_fail += 1
            continue

        if apply_new_geometry(element, new_line):
            count_ok += 1
        else:
            count_fail += 1

# ─────────────────────────────────────────────────────────────
# Summary
# ─────────────────────────────────────────────────────────────

msg_parts = [
    "Chamfer Lines 執行完成！",
    "",
    "偵測閾值   : {} mm".format(threshold_mm),
    "找到端點對 : {} 組".format(len(pairs)),
    "─────────────────────",
    "✓ 成功延伸 : {} 條".format(count_ok),
]
if count_skip > 0:
    msg_parts.append("⚠ 跳過退化  : {} 條".format(count_skip))
if skipped_parallel > 0:
    msg_parts.append("⚠ 平行跳過  : {} 對".format(skipped_parallel))
if count_fail > 0:
    msg_parts.append("✗ 失敗      : {} 條 (請查看輸出視窗)".format(count_fail))

forms.alert("\n".join(msg_parts), title="ChamferLines")
