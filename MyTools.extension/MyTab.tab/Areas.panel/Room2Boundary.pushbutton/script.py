# -*- coding: utf-8 -*-
"""
Recreates a Dynamo flow specifically for Extracting Room Boundaries
and Dimensioning them in a targeted View.

Features:
- Select Level
- Select View (Area Plan or Floor Plan)
- Select TextNoteType
- Options to create DetailCurves or AreaBoundaryLines
- Option to place lengths
- Options to delete old dimensions
"""

import sys
from pyrevit import forms, revit, script
from pyrevit import DB

doc = revit.doc

def feet_to_meter(feet):
    return round(feet * 0.3048, 2)

# --- 1. Select Level ---
levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
if not levels:
    forms.alert("No Levels found in the model.")
    script.exit()

selected_level = forms.SelectFromList.show(
    levels,
    name_attr='Name',
    title="1. 選擇來源 Level (依據此樓層獲取 Rooms)",
    button_name="下一步"
)
if not selected_level:
    script.exit()

# --- 2. Select Target View ---
# We generally want Area Plans, but Floor Plans are OK too
all_views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
valid_views = []
for v in all_views:
    if v.IsTemplate: continue
    if v.ViewType == DB.ViewType.AreaPlan or v.ViewType == DB.ViewType.FloorPlan:
        valid_views.append(v)

if not valid_views:
    forms.alert("No AreaPlan or FloorPlan views found in the model.")
    script.exit()

# To help users, format view name with its type
class ViewItem(forms.TemplateListItem):
    @property
    def name(self):
        return "[{}] {}".format(self.item.ViewType, self.item.Name)

view_items = [ViewItem(v) for v in valid_views]
selected_view_item = forms.SelectFromList.show(
    view_items,
    title="2. 選擇目標 View (優先選擇 Area Plan)",
    button_name="下一步"
)
if not selected_view_item:
    script.exit()
selected_view = getattr(selected_view_item, 'item', selected_view_item)

# --- 3. Select TextNoteType ---
text_types = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).ToElements()
if not text_types:
    forms.alert("No TextNoteTypes found.")
    script.exit()

# To help users, format text type name safely
class TextTypeItem(forms.TemplateListItem):
    @property
    def name(self):
        try:
            return self.item.Name
        except Exception:
            try:
                # Fallback to BuiltInParameter if .Name fails in some Revit versions
                param = self.item.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    return param.AsString()
                return "Unknown Type [{}]".format(self.item.Id.IntegerValue)
            except:
                return "Unknown Type [{}]".format(self.item.Id.IntegerValue)

text_type_items = [TextTypeItem(t) for t in text_types]
selected_text_type_item = forms.SelectFromList.show(
    text_type_items,
    title="3. 選擇文字樣式 (TextNoteType)",
    button_name="下一步"
)
if not selected_text_type_item:
    script.exit()
selected_text_type = getattr(selected_text_type_item, 'item', selected_text_type_item)

# --- 4. Toggle Options ---
opts = [
    "建立 DetailCurves (適合視覺標註)",
    "建立 Area Boundary Lines (適合面積計算，需為 Area Plan)",
    "建立長度標註 (TextNotes)",
    "執行前先刪除目標視圖中舊標註 (與所選樣式相同的 TextNotes)",
    "避免重複建立標註 (若該處已有數字不重複建立)"
]
selected_opts = forms.SelectFromList.show(
    opts, 
    title="4. 選擇執行選項 (可複選)", 
    multiselect=True,
    button_name="執行"
)
if not selected_opts:
    script.exit()

create_detail_curves = "建立 DetailCurves (適合視覺標註)" in selected_opts
create_area_lines = "建立 Area Boundary Lines (適合面積計算，需為 Area Plan)" in selected_opts
create_dimensions = "建立長度標註 (TextNotes)" in selected_opts
delete_old = "執行前先刪除目標視圖中舊標註 (與所選樣式相同的 TextNotes)" in selected_opts
avoid_duplicates = "避免重複建立標註 (若該處已有數字不重複建立)" in selected_opts

# Validate Area Plan requirement
if create_area_lines and selected_view.ViewType != DB.ViewType.AreaPlan:
    forms.alert("所選視圖不是 Area Plan，無法建立 Area Boundary Lines。\n將取消此選項。")
    create_area_lines = False

# --- Prepare Data ---
rooms = DB.FilteredElementCollector(doc)\
          .OfCategory(DB.BuiltInCategory.OST_Rooms)\
          .WhereElementIsNotElementType()\
          .ToElements()

target_rooms = [r for r in rooms if r.LevelId == selected_level.Id and r.Area > 0]
if not target_rooms:
    forms.alert("在選定樓層 {} 找不到任何有效的 Rooms。".format(selected_level.Name))
    script.exit()

# Setup boundary options
b_options = DB.SpatialElementBoundaryOptions()
b_options.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish

# Fetch existing TextNotes in the view to avoid duplicates
existing_texts = DB.FilteredElementCollector(doc, selected_view.Id)\
                   .OfClass(DB.TextNote)\
                   .ToElements()

# --- Execution ---
count_lines = 0
count_area_lines = 0
count_dims = 0
count_skipped_curves = 0
count_deleted_old = 0
count_skipped_duplicates = 0

with revit.Transaction("Room Boundaries & Dimensions"):
    
    # Optional: Delete old dimensions in this view
    if delete_old:
        for tn in existing_texts:
            if tn.GetTypeId() == selected_text_type.Id:
                doc.Delete(tn.Id)
                count_deleted_old += 1
        # Refresh existing_texts since we just deleted some
        existing_texts = DB.FilteredElementCollector(doc, selected_view.Id)\
                           .OfClass(DB.TextNote)\
                           .ToElements()

    # Pre-calculate what text values and positions are already there
    existing_infos = []
    if avoid_duplicates and existing_texts:
        for tn in existing_texts:
            existing_infos.append((tn.Coord, tn.Text))

    for room in target_rooms:
        boundaries = room.GetBoundarySegments(b_options)
        if not boundaries:
            continue
            
        for segment_list in boundaries:
            for segment in segment_list:
                curve = segment.GetCurve()
                
                # Check curve type (we only handle lines according to use case, but arc is possible for room bounds)
                if not isinstance(curve, DB.Line):
                    count_skipped_curves += 1
                    # Can be handled later if requested
                    continue
                
                # We have a Line
                p1 = curve.GetEndPoint(0)
                p2 = curve.GetEndPoint(1)
                
                # For 2D drawing in views, it's safer to flatten them to the view's Z
                # But DetailCurve/AreaBoundary requires exact plane of the view if needed
                
                # Create DetailCurves
                if create_detail_curves:
                    try:
                        doc.Create.NewDetailCurve(selected_view, curve)
                        count_lines += 1
                    except Exception as e:
                        print("DetailCurve creation failed: {}".format(e))
                
                # Create Area Boundary Lines
                if create_area_lines:
                    try:
                        # Area boundary line needs a SketchPlane in the view
                        sp = selected_view.SketchPlane
                        if not sp:
                            if hasattr(selected_view, 'GenLevel') and selected_view.GenLevel:
                                sp = DB.SketchPlane.Create(doc, selected_view.GenLevel.Id)
                                selected_view.SketchPlane = sp
                            else:
                                raise Exception("No GenLevel to create SketchPlane.")
                                
                        doc.Create.NewAreaBoundaryLine(sp, curve, selected_view)
                        count_area_lines += 1
                    except Exception as e:
                        print("Area Boundary Line creation failed: {}".format(e))
                
                # Create Dimensions
                if create_dimensions:
                    mid_x = (p1.X + p2.X) / 2
                    mid_y = (p1.Y + p2.Y) / 2
                    # Align to view's elevation just in case
                    mid_z = selected_level.ProjectElevation 
                    if hasattr(selected_view, 'GenLevel') and selected_view.GenLevel:
                        mid_z = selected_view.GenLevel.ProjectElevation
                        
                    mid_pt = DB.XYZ(mid_x, mid_y, mid_z)
                    
                    length_m = feet_to_meter(p1.DistanceTo(p2))
                    text_str = str(length_m)
                    
                    # Duplicate check
                    is_dup = False
                    if avoid_duplicates:
                        # 2.0 feet tolerance (~60cm) for finding same number text nearby
                        # This covers the thickness of most walls between adjacent rooms
                        tolerance = 2.0
                        for ex_coord, ex_text in existing_infos:
                            if ex_text == text_str and mid_pt.DistanceTo(ex_coord) < tolerance:
                                is_dup = True
                                break
                                
                    if is_dup:
                        count_skipped_duplicates += 1
                    else:
                        try:
                            tn_options = DB.TextNoteOptions()
                            tn_options.TypeId = selected_text_type.Id
                            tn_options.HorizontalAlignment = DB.HorizontalTextAlignment.Center
                            tn_options.VerticalAlignment = DB.VerticalTextAlignment.Middle
                            
                            DB.TextNote.Create(doc, selected_view.Id, mid_pt, text_str, tn_options)
                            
                            # Add to existing_infos so we don't duplicate within same loop
                            if avoid_duplicates:
                                existing_infos.append((mid_pt, text_str))
                                
                            count_dims += 1
                        except Exception as e:
                            print("TextNote creation failed: {}".format(e))

# --- Summary ---
msg = "執行摘要:\n"
msg += "- 來源樓層: {}\n".format(selected_level.Name)
msg += "- 目標視圖: {}\n".format(selected_view.Name)
msg += "--------------------------------------\n"
if create_detail_curves:
    msg += "建立 DetailCurves: {} 條\n".format(count_lines)
if create_area_lines:
    msg += "建立 Area Boundary Lines: {} 條\n".format(count_area_lines)
if create_dimensions:
    msg += "建立長度標註: {} 個\n".format(count_dims)
msg += "--------------------------------------\n"
msg += "跳過非直線曲線 (Arc等): {} 條\n".format(count_skipped_curves)

if delete_old:
    msg += "以刪除舊標註: {} 個\n".format(count_deleted_old)
if avoid_duplicates:
    msg += "因避免重複而跳過標註: {} 個".format(count_skipped_duplicates)

forms.alert(msg, title="執行完成")
