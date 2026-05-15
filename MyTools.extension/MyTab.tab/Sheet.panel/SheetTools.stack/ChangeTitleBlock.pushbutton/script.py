# -*- coding: utf-8 -*-
"""批次更換所選圖紙的圖框 (TitleBlock)。

選擇圖紙 → 選擇目標圖框類型 → 一次替換所有選中圖紙的 TitleBlock。
使用 ChangeTypeId() 保留圖框位置與共用參數值。
"""

from pyrevit import forms, revit, script, DB

doc = revit.doc


# --- 1. 收集專案中所有 TitleBlock Types ---
tb_symbols = (
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
    .WhereElementIsElementType()
    .ToElements()
)

if not tb_symbols:
    forms.alert("專案中沒有已載入的圖框族群 (TitleBlock Family)。", exitscript=True)


# --- 2. 選擇圖紙 ---
selected_sheets = forms.select_sheets(
    button_name="選擇圖紙",
    use_selection=True,
    include_placeholder=False,
)
if not selected_sheets:
    script.exit()


# --- 3. 顯示目前圖紙上的圖框資訊，讓使用者確認 ---
print("=" * 60)
print("已選取 {} 張圖紙，目前圖框狀態：".format(len(selected_sheets)))
print("-" * 60)
for sheet in selected_sheets:
    tblocks = revit.query.get_sheet_tblocks(sheet)
    if tblocks:
        for tb in tblocks:
            tb_type = doc.GetElement(tb.GetTypeId())
            family_name = tb_type.FamilyName if tb_type else "N/A"
            type_name = tb_type.get_Parameter(
                DB.BuiltInParameter.SYMBOL_NAME_PARAM
            ).AsString() if tb_type else "N/A"
            print(
                "  {0} - {1}  →  圖框: {2} : {3}".format(
                    sheet.SheetNumber, sheet.Name, family_name, type_name
                )
            )
    else:
        print(
            "  {0} - {1}  →  (無圖框)".format(sheet.SheetNumber, sheet.Name)
        )
print("=" * 60)


# --- 4. 選擇目標 TitleBlock Type ---
class TitleBlockItem(forms.TemplateListItem):
    """顯示 FamilyName : TypeName 格式"""
    @property
    def name(self):
        try:
            family_name = self.item.FamilyName
            param = self.item.get_Parameter(
                DB.BuiltInParameter.SYMBOL_NAME_PARAM
            )
            type_name = param.AsString() if param else "?"
            return "{} : {}".format(family_name, type_name)
        except Exception:
            return "Unknown [{}]".format(self.item.Id.IntegerValue)


tb_items = sorted(
    [TitleBlockItem(s) for s in tb_symbols],
    key=lambda x: x.name,
)

selected_tb_item = forms.SelectFromList.show(
    tb_items,
    title="選擇目標圖框類型 (TitleBlock Type)",
    button_name="更換圖框",
)
if not selected_tb_item:
    script.exit()

target_tb_type = selected_tb_item  # TemplateListItem wraps .item


# --- 5. 批次更換 ---
count_success = 0
count_skipped = 0
count_no_tb = 0
count_error = 0
errors = []

with revit.Transaction("Batch Change TitleBlock"):
    for sheet in selected_sheets:
        tblocks = revit.query.get_sheet_tblocks(sheet)
        if not tblocks:
            count_no_tb += 1
            continue

        for tb in tblocks:
            current_type_id = tb.GetTypeId()
            if current_type_id == target_tb_type.Id:
                count_skipped += 1
                continue

            try:
                tb.ChangeTypeId(target_tb_type.Id)
                count_success += 1
            except Exception as e:
                count_error += 1
                errors.append(
                    "{} - {}: {}".format(
                        sheet.SheetNumber, sheet.Name, str(e)
                    )
                )


# --- 6. 執行摘要 ---
msg = "批次更換圖框完成！\n"
msg += "=" * 40 + "\n"
msg += "成功更換: {} 張\n".format(count_success)
if count_skipped:
    msg += "已是目標圖框 (跳過): {} 張\n".format(count_skipped)
if count_no_tb:
    msg += "無圖框的圖紙: {} 張\n".format(count_no_tb)
if count_error:
    msg += "更換失敗: {} 張\n".format(count_error)
    msg += "-" * 40 + "\n"
    for err in errors:
        msg += "  ⚠ {}\n".format(err)

forms.alert(msg, title="更換圖框結果")
