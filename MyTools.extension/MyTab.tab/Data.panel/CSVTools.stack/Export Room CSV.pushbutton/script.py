# -*- coding: utf-8 -*-
from __future__ import print_function
from pyrevit import forms, revit, DB
import csv
import io

doc = revit.doc

# ---------- 文字/編碼與正規化 ----------
def to_unicode(x):
    if isinstance(x, bytes):
        try:
            return x.decode('utf-8')
        except:
            try:
                return x.decode('utf-8-sig')
            except:
                return x.decode('cp950', 'ignore')
    return x

def normalize(s):
    if s is None:
        return u''
    s = to_unicode(s)
    s = s.replace(u'\u3000', u' ')  # 全形空白
    # 全形數字/字母 -> 半形
    fw = u'０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺabcdefghijklmnopqrstuvwxyz'
    hw = u'0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
    tbl = dict((ord(fw[i]), hw[i]) for i in range(len(hw)))
    return s.translate(tbl).strip()

def get_param_text(param):
    if not param or not param.HasValue:
        return u''
    st = param.StorageType
    try:
        if st == DB.StorageType.String:
            return param.AsString()
        elif st == DB.StorageType.Integer:
            # Check if it's a Yes/No parameter or similar
            val_str = param.AsValueString()
            if val_str:
                return val_str
            return unicode(param.AsInteger())
        elif st == DB.StorageType.Double:
            # AsValueString usually gets the correct formatted string with units
            val_str = param.AsValueString()
            if val_str is not None and val_str != "":
                # If we just want the number without unit suffix, we can parse it, 
                # but AsValueString is safest to get what the user sees in Revit.
                # However, for CSV export, sometimes we want just the number.
                # Let's try to get just the number in the correct display unit.
                
                # In newer Revit APIs (2021+), we use GetUnitTypeId
                try:
                    unit_id = param.Definition.GetUnitTypeId()
                    # Format options
                    if str(unit_id.TypeId) != "":
                        # Better just return the AsValueString and let user see units
                        return val_str
                except:
                    pass
                return val_str
            else:
                # Fallback if AsValueString fails for some reason
                return unicode(param.AsDouble())
        else:
            return param.AsValueString()
    except:
        return u''

def set_param_value(param, val):
    # This is for the export script, but I might as well define it in case we need it
    pass

def main():
    # 1. 收集所有房間
    rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
    if not rooms:
        forms.alert(u'目前的模型中找不到任何房間。', exitscript=True)
        return
        
    #過濾掉未放置(Unplaced)的房間
    placed_rooms = [r for r in rooms if r.Location is not None and r.Area > 0]
    
    # 2. 收集所有可能的參數名稱
    param_names = set()
    for room in placed_rooms[:50]: # 只取前50個來找參數，通常已經夠了
        for p in room.Parameters:
            if p.Definition:
                param_names.add(p.Definition.Name)
                
    param_names = sorted(list(param_names))
    
    # 常用的預設勾選參數
    default_checked = ['Number', 'Name', 'Department', 'Comments', 'Base Finish', 'Ceiling Finish', 'Wall Finish', 'Floor Finish', 'Area', 'Level']
    
    # 讓使用者選擇要匯出的欄位
    class ParamOption(forms.TemplateListItem):
        @property
        def name(self):
            return self.item
            
    options = [ParamOption(p, checked=(p in default_checked or p == u'編號' or p == u'名稱')) for p in param_names]
    
    selected_params = forms.SelectFromList.show(options, multiselect=True, title=u'選取要匯出的房間參數', button_name=u'下一步')
    
    if not selected_params:
        forms.alert(u'未選取任何參數，取消匯出。', exitscript=True)
        return
        
    # 強制將 Number (或編號) 加到第一欄，作為日後匯入的 Key
    key_params = ['Number', u'編號', 'Room Number', u'房間編號']
    actual_key = None
    for kp in key_params:
        if kp in param_names:
            actual_key = kp
            break
            
    if actual_key and actual_key not in selected_params:
        selected_params.insert(0, actual_key)
    elif actual_key and actual_key in selected_params:
        selected_params.remove(actual_key)
        selected_params.insert(0, actual_key)

    # 3. 匯出邏輯
    csv_path = forms.save_file(file_ext='csv', title=u'儲存房間資料 CSV')
    if not csv_path:
        forms.alert(u'已取消儲存。', exitscript=True)
        return
        
    try:
        exported_count = 0
        with io.open(csv_path, 'wb') as f:
            # 寫入 BOM 使 Excel 可以直接閱讀 UTF-8
            f.write(b'\xef\xbb\xbf')
            
            # 使用 unicodecsv 或標準 csv 並編碼
            writer = csv.writer(f)
            
            # 表頭
            header = [p.encode('utf-8') for p in selected_params]
            writer.writerow(header)
            
            # 資料
            for room in placed_rooms:
                row_data = []
                for p_name in selected_params:
                    # 處理 Level 特殊情況 (它是獨立的屬性)
                    if p_name == 'Level' or p_name == u'樓層':
                        level = room.Level
                        val = level.Name if level else u''
                    else:
                        p = room.LookupParameter(p_name)
                        val = get_param_text(p) if p else u''
                    
                    row_data.append(val.encode('utf-8', 'ignore'))
                
                writer.writerow(row_data)
                exported_count += 1
                
        forms.alert(u'成功匯出 {} 個房間的資料至：\n{}'.format(exported_count, csv_path), title=u'匯出完成')
        
    except Exception as e:
        forms.alert(u'匯出時發生錯誤：\n{}'.format(e), exitscript=True)

if __name__ == '__main__':
    main()
