# -*- coding: utf-8 -*-
"""
Recreates Room elements on target levels at the exact same XY coordinates
as the selected rooms from a reference level.
Also copies Name, Number (with renumbering options), Department, and Comments.
"""

import sys
import re
from pyrevit import forms, revit, script
from pyrevit import DB

doc = revit.doc

# 1. Select Reference Level (來源樓層)
levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
if not levels:
    forms.alert("專案中沒有找到任何樓層 (Levels)。")
    script.exit()

# Sort levels by project elevation for easier selection
levels = sorted(levels, key=lambda l: l.ProjectElevation)

selected_source_level = forms.SelectFromList.show(
    levels,
    name_attr='Name',
    title="1. 選擇參考來源樓層",
    button_name="下一步"
)
if not selected_source_level:
    script.exit()

# 2. Get Rooms from Source Level
rooms = DB.FilteredElementCollector(doc)\
          .OfCategory(DB.BuiltInCategory.OST_Rooms)\
          .WhereElementIsNotElementType()\
          .ToElements()

# Filter active placed rooms on source level
source_rooms = [r for r in rooms if r.LevelId == selected_source_level.Id and r.Area > 0]
if not source_rooms:
    forms.alert("在來源樓層 [{}] 中找不到任何已放置的房間 (Rooms)。".format(selected_source_level.Name))
    script.exit()

# Wrapper class for room selection
class RoomListItem(forms.TemplateListItem):
    @property
    def name(self):
        number = self.item.Number
        name_val = self.item.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
        return "[{}] {}".format(number, name_val)

room_items = [RoomListItem(r) for r in source_rooms]
room_items = sorted(room_items, key=lambda ri: ri.item.Number)

selected_rooms = forms.SelectFromList.show(
    room_items,
    title="2. 選擇要複製的房間 (可複選，預設全選)",
    multiselect=True,
    default_selected=room_items,
    button_name="下一步"
)
if not selected_rooms:
    script.exit()

# 3. Select Target Levels (目標樓層)
target_level_candidates = [l for l in levels if l.Id != selected_source_level.Id]
if not target_level_candidates:
    forms.alert("沒有其他可選的目標樓層。")
    script.exit()

selected_targets = forms.SelectFromList.show(
    target_level_candidates,
    name_attr='Name',
    title="3. 選擇要放置新房間的目標樓層 (可複選)",
    multiselect=True,
    button_name="下一步"
)
if not selected_targets:
    script.exit()

# 4. Select Numbering Option
opts = [
    "智能替換首位數字 (例如 101 -> 在 2F 為 201, 3F 為 301)",
    "保持相同房號 (可能產生房號重複警告)",
    "加上目標樓層前綴 (例如 101 -> 2F-101)",
]
selected_opt = forms.SelectFromList.show(
    opts,
    title="4. 選擇新房間的房號命名原則",
    button_name="開始執行"
)
if not selected_opt:
    script.exit()

# Helper function to parse digit prefix from Level Name
def get_level_number_prefix(level_name):
    # Search for first digit sequence (supports negative numbers, e.g. B1 -> 1, -1F -> -1, 2F -> 2)
    match = re.search(r'-?\d+', level_name)
    if match:
        return match.group(0)
    return ""

def generate_new_number(orig_num, target_level_name, mode):
    if mode == "智能替換首位數字 (例如 101 -> 在 2F 為 201, 3F 為 301)":
        target_prefix = get_level_number_prefix(target_level_name)
        if target_prefix and orig_num and len(orig_num) > 0:
            # Replace the first character if it is a digit
            if orig_num[0].isdigit():
                return target_prefix + orig_num[1:]
            else:
                return target_prefix + "_" + orig_num
        return orig_num
    elif mode == "加上目標樓層前綴 (例如 101 -> 2F-101)":
        return "{}-{}".format(target_level_name, orig_num)
    else:
        # Keep same
        return orig_num

# 5. Place Rooms
success_count = 0
failed_count = 0
fail_details = []

with revit.Transaction("對齊放置房間 (Place Rooms By Reference)"):
    for t_level in selected_targets:
        t_level_name = t_level.Name
        
        for s_room in selected_rooms:
            loc = s_room.Location
            if not loc or not isinstance(loc, DB.LocationPoint):
                failed_count += 1
                fail_details.append("房間 [{}] 沒有有效的 LocationPoint。".format(s_room.Number))
                continue
            
            pt = loc.Point
            uv_pt = DB.UV(pt.X, pt.Y)
            
            try:
                # Place new room at target level
                new_room = doc.Create.NewRoom(t_level, uv_pt)
                
                if new_room is None:
                    failed_count += 1
                    fail_details.append("樓層 [{}]: 坐標 ({:.2f}, {:.2f}) 無法放置房間 (可能該處非封閉區域)。".format(
                        t_level_name, pt.X, pt.Y
                    ))
                    continue
                
                # Copy Room Name
                s_name = s_room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
                new_room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).Set(s_name)
                
                # Set Room Number
                new_num = generate_new_number(s_room.Number, t_level_name, selected_opt)
                new_room.Number = new_num
                
                # Copy standard parameters if they exist
                # Comments
                s_comments = s_room.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString()
                if s_comments:
                    new_room.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(s_comments)
                
                # Department
                s_dept = s_room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString()
                if s_dept:
                    new_room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).Set(s_dept)
                
                success_count += 1
                
            except Exception as e:
                failed_count += 1
                fail_details.append("樓層 [{}]: 房間 [{}] 放置或設定參數失敗: {}".format(
                    t_level_name, s_room.Number, str(e)
                ))

# 6. Report Summary
report = "執行結果摘要：\n"
report += "--------------------------------------\n"
report += "- 來源樓層: {}\n".format(selected_source_level.Name)
report += "- 成功放置房間: {} 個\n".format(success_count)
report += "- 失敗或跳過: {} 個\n".format(failed_count)
report += "--------------------------------------\n"

if fail_details:
    report += "\n詳細錯誤/跳過原因 (前 15 筆)：\n"
    for detail in fail_details[:15]:
        report += "- {}\n".format(detail)
    if len(fail_details) > 15:
        report += "- ...以及其他 {} 筆錯誤\n".format(len(fail_details) - 15)

forms.alert(report, title="對齊放置房間已完成")
