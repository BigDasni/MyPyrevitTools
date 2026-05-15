# -*- coding: utf-8 -*-
"""
SyncViewportPos - 同步圖紙視圖位置
以某張圖紙上的視圖為基準，批次將其位置套用到其他圖紙的對應視圖上。
支援關鍵字搜尋視圖名稱。
"""
from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ObjectType

doc = revit.doc
uidoc = revit.uidoc

# ─────────────────────────────────────────
# 前置檢查：確認目前在 Sheet 視窗中
# ─────────────────────────────────────────
active_view = uidoc.ActiveView
if not isinstance(active_view, DB.ViewSheet):
    forms.alert(
        "請先開啟一張圖紙（Sheet）視窗，再執行此工具。\n\n"
        "目前的視窗類型：{}\n\n"
        "操作方式：在 Project Browser 中雙擊任一張圖紙開啟，再重新執行工具。".format(
            str(active_view.ViewType).split(".")[-1]
        ),
        title="請先開啟 Sheet 視圖"
    )
    script.exit()

# ─────────────────────────────────────────
# Step 1：使用者點選基準 Viewport
# 說明文字顯示在 Revit 底部狀態列，不彈出 alert
# （alert 會讓 Revit 失去焦點，導致 PickObject 立刻被取消）
# ─────────────────────────────────────────
try:
    ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        "【同步視圖位置】在圖紙上點選基準視圖框  |  Esc = 取消"
    )
except Exception as pick_err:
    err_msg = str(pick_err)
    # Revit 按 Esc 取消時訊息包含 "cancelled" / "canceled"，靜默退出即可
    if "cancel" in err_msg.lower() or "esc" in err_msg.lower() or not err_msg:
        script.exit()
    forms.alert(
        "點選視圖時發生錯誤：\n{}\n\n"
        "請確認：目前正在圖紙（Sheet）視窗中操作。".format(err_msg),
        title="錯誤"
    )
    script.exit()

base_elem = doc.GetElement(ref.ElementId)

if not isinstance(base_elem, DB.Viewport):
    forms.alert(
        "選取的不是 Viewport！\n\n"
        "請確認：\n"
        "• 點選的是視圖框的邊界線，而非視圖內部\n"
        "• 目前在圖紙（Sheet）視窗中操作",
        title="選取錯誤，請重新執行工具"
    )
    script.exit()

base_vp = base_elem
base_center = base_vp.GetBoxCenter()
base_view = doc.GetElement(base_vp.ViewId)
base_view_name = base_view.Name
base_sheet_id = base_vp.SheetId
base_sheet = doc.GetElement(base_sheet_id)

# ─────────────────────────────────────────
# Step 2：輸入搜尋關鍵字（過濾目標視圖名稱）
# ─────────────────────────────────────────
keyword = forms.ask_for_string(
    prompt=(
        "輸入關鍵字以搜尋目標視圖名稱\n"
        "（例如：plan、section、elevation）\n\n"
        "基準視圖：{}\n"
        "基準圖紙：{} - {}".format(
            base_view_name,
            base_sheet.SheetNumber,
            base_sheet.Name
        )
    ),
    title="搜尋視圖關鍵字",
    default=base_view_name
)

if keyword is None:
    script.exit()

keyword = keyword.strip()
if not keyword:
    forms.alert("關鍵字不能為空！", title="錯誤")
    script.exit()

# ─────────────────────────────────────────
# Step 3：選擇目標圖紙（多選）
# ─────────────────────────────────────────
all_sheets = (
    DB.FilteredElementCollector(doc)
    .OfClass(DB.ViewSheet)
    .ToElements()
)

# 排除基準視圖所在圖紙，並依圖號排序
target_sheets = sorted(
    [s for s in all_sheets if s.Id != base_sheet_id],
    key=lambda s: s.SheetNumber
)

if not target_sheets:
    forms.alert("找不到其他圖紙可套用。", title="無目標圖紙")
    script.exit()

sheet_map = {
    "{} - {}".format(s.SheetNumber, s.Name): s
    for s in target_sheets
}

selected_keys = forms.SelectFromList.show(
    sorted(sheet_map.keys()),
    title="選擇要套用的目標圖紙（可多選）",
    multiselect=True,
    button_name="套用位置"
)

if not selected_keys:
    script.exit()

# ─────────────────────────────────────────
# Step 4：在每張圖紙中，依關鍵字過濾並讓使用者確認匹配的 Viewport
# ─────────────────────────────────────────
success_list = []
skipped_list = []
multi_match_list = []

def get_viewport_label(vp):
    """取得 Viewport 的顯示標籤（視圖名稱 + ViewType）"""
    v = doc.GetElement(vp.ViewId)
    if v:
        return "{} [{}]".format(v.Name, str(v.ViewType).split(".")[-1])
    return "（未知視圖）"

with revit.Transaction("同步視圖位置"):
    for key in selected_keys:
        sheet = sheet_map[key]

        # 取得此圖紙上所有 Viewport
        vps = list(
            DB.FilteredElementCollector(doc, sheet.Id)
            .OfClass(DB.Viewport)
            .ToElements()
        )

        # 依關鍵字過濾（不區分大小寫）
        kw_lower = keyword.lower()
        matched_vps = []
        for vp in vps:
            v = doc.GetElement(vp.ViewId)
            if v and kw_lower in v.Name.lower():
                matched_vps.append(vp)

        if not matched_vps:
            skipped_list.append("{} → 找不到包含「{}」的視圖".format(key, keyword))
            continue

        # 若只有一個符合，直接套用
        if len(matched_vps) == 1:
            target_vp = matched_vps[0]
        else:
            # 多個符合時，讓使用者選擇要套用哪一個
            vp_label_map = {get_viewport_label(vp): vp for vp in matched_vps}
            chosen = forms.SelectFromList.show(
                sorted(vp_label_map.keys()),
                title="【{}】找到多個符合視圖，請選擇要套用的".format(key),
                multiselect=False,
                button_name="選擇此視圖"
            )
            if not chosen:
                skipped_list.append("{} → 使用者略過".format(key))
                continue
            target_vp = vp_label_map[chosen]

        # 套用位置（只同步 XY，保留 Z 不變）
        old_center = target_vp.GetBoxCenter()
        new_center = DB.XYZ(base_center.X, base_center.Y, old_center.Z)
        target_vp.SetBoxCenter(new_center)

        target_view = doc.GetElement(target_vp.ViewId)
        success_list.append(
            "{} → 視圖：{}".format(key, target_view.Name if target_view else "?")
        )

# ─────────────────────────────────────────
# Step 5：顯示結果報告
# ─────────────────────────────────────────
report_lines = []

if success_list:
    report_lines.append("✅ 成功套用（{}張）：".format(len(success_list)))
    for item in success_list:
        report_lines.append("  • " + item)

if skipped_list:
    report_lines.append("")
    report_lines.append("⚠️ 略過（{}張）：".format(len(skipped_list)))
    for item in skipped_list:
        report_lines.append("  • " + item)

forms.alert(
    "\n".join(report_lines) if report_lines else "沒有任何操作被執行。",
    title="同步視圖位置 — 完成報告"
)
