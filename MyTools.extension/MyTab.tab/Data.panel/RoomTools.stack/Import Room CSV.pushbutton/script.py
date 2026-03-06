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

# ---------- 讀 CSV ----------
def read_csv_any(path):
    encodings = ('utf-8-sig', 'utf-8', 'cp950')
    last_err = None
    for enc in encodings:
        try:
            rows = []
            with open(path, 'rb') as f:
                data = f.read()
                text = data.decode(enc, 'ignore')
                reader = csv.reader(text.splitlines())
                headers = [normalize(h) if h else u'' for h in next(reader)]
                # 保證每個欄位都有名稱
                for i, h in enumerate(headers):
                    if not h:
                        headers[i] = u'Column{}'.format(i+1)
                for row in reader:
                    item = {}
                    for i, h in enumerate(headers):
                        val = row[i] if i < len(row) else u''
                        item[h] = normalize(val)
                    rows.append(item)
            return headers, rows, enc
        except StopIteration:
            # 空檔案或沒有表頭
            return [], [], enc
        except Exception as e:
            last_err = e
            continue
    raise Exception(u'CSV 讀取失敗（嘗試 utf-8-sig/utf-8/cp950）。最後錯誤：{}'.format(last_err))

# ---------- Revit 取值/寫值 ----------
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
                try:
                    unit_id = param.Definition.GetUnitTypeId()
                    if str(unit_id.TypeId) != "":
                        return val_str
                except:
                    pass
                return val_str
            else:
                return unicode(param.AsDouble())
        else:
            return param.AsValueString()
    except:
        return u''

def set_param_value(param, val):
    # 如果是唯讀參數，例如面積等，不能寫入
    if param.IsReadOnly:
        raise Exception(u'參數為唯讀 (Read-Only)')
        
    st = param.StorageType
    if st == DB.StorageType.String:
        param.Set(val)
    elif st == DB.StorageType.Integer:
        if not val:
            # 處理清空整數的情況
            pass 
        else:
            param.Set(int(val))
    elif st == DB.StorageType.Double:
        if not val:
            pass
        else:
            # Try to parse string with units back to internal Double
            try:
                # Revit 2021+ provides UnitUtils.TryParse
                unit_id = param.Definition.GetUnitTypeId()
                parsed_val = doc.Application.Create.NewProjectDocument(DB.UnitSystem.Metric).GetUnits()
                # Use SetValueString which handles built-in unit conversion nicely
                success = param.SetValueString(val)
                if not success:
                    # Fallback to direct float cast if SetValueString fails (e.g., pure number string)
                    param.Set(float(val))
            except Exception:
                # Fallback for older versions or if GetUnitTypeId fails
                try:
                    success = param.SetValueString(val)
                    if not success:
                        param.Set(float(val))
                except:
                    param.Set(float(val))
    else:
        # 其他型別 (例如 ElementId)，直接用字串嘗試寫入通常會失敗
        param.Set(val)

# ---------- 主流程 ----------
def main():
    csv_path = forms.pick_file(file_ext='csv', title=u'選擇要匯入的房間資料 CSV (建議 UTF-8)')
    if not csv_path:
        forms.alert(u'已取消。')
        return

    try:
        headers, rows, used_enc = read_csv_any(csv_path)
    except Exception as e:
        forms.alert(u'讀取 CSV 失敗：{}'.format(e), exitscript=True)
        return

    if not headers or not rows:
        forms.alert(u'CSV 無資料或無表頭。', exitscript=True)
        return

    # 1) CSV 裡的對應鍵（預設找 Number 或 編號）
    default_key_col = u'Number'
    if u'編號' in headers: default_key_col = u'編號'
    elif u'Room Number' in headers: default_key_col = u'Room Number'
    elif 'Number' in headers: default_key_col = 'Number'
    elif headers: default_key_col = headers[0]
    
    key_col = forms.ask_for_one_item(headers, title=u'① 選擇 CSV 中作為比對索引的欄位 (通常是「編號/Number」)', default=default_key_col)
    if not key_col:
        forms.alert(u'未選擇對應鍵。', exitscript=True); return

    # 2) Revit 參數名 (預設 Number)
    revit_key_param = forms.ask_for_string(
        default='Number',
        prompt=u'② 輸入 Revit 「房間(Room)」用來比對的參數名（預設為「Number」或「編號」）。'
    )
    if not revit_key_param:
        forms.alert(u'未提供 Revit 參數名。', exitscript=True); return

    # 3) 選擇要回寫哪些欄位
    candidates = [h for h in headers if h != key_col]
    if not candidates:
        forms.alert(u'CSV 內沒有其他可用來更新的欄位 (只有索引欄)。', exitscript=True); return
        
    title_msg = u'③ 選擇要回寫的欄位（可多選）\n---\n注意：匯出的唯讀欄位(如 Area/面積)無法寫回 Revit。'
    target_cols = forms.SelectFromList.show(candidates, multiselect=True, title=title_msg, button_name=u'執行匯入')

    if not target_cols:
        forms.alert(u'未選擇回寫欄位。', exitscript=True); return

    # 4) 整理現有房間索引
    forms.alert(u'建立房間索引：依 [{}] 比對中...'.format(revit_key_param), warn_icon=False)
    
    # 建立 Room Index Dictionary { "101": RoomElement, ... }
    room_index = {}
    for r in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType():
        if not r: continue
        p = r.LookupParameter(revit_key_param)
        val = normalize(get_param_text(p))
        if val:
            room_index.setdefault(val, r)

    if not room_index:
        forms.alert(u'警告：在目前的 Revit 專案中，找不到任何帶有參數「{}」的房間。匯入將強制停止。'.format(revit_key_param), exitscript=True); return

    # 5) 寫入準備
    param_alias = {}
    updated_rooms = 0
    notfound_in_revit = 0
    skipped_rows = 0
    failures = []

    t = DB.Transaction(doc, u'CSV 回寫房間參數')
    t.Start()
    try:
        for row in rows:
            key_val = normalize(row.get(key_col))
            if not key_val:
                skipped_rows += 1
                continue

            room = room_index.get(key_val)
            if not room:
                notfound_in_revit += 1
                continue

            # 此房間需要修改嗎？
            room_updated = False
            
            for col in target_cols:
                raw_val = row.get(col)
                if raw_val is None:
                    continue
                val = normalize(raw_val)

                # 參數映射 (如果 CSV 表頭和 Revit 參數不同，第一次遇到會問)
                revit_param_name = param_alias.get(col)
                if not revit_param_name:
                    revit_param_name = col
                    p_room = room.LookupParameter(revit_param_name)
                    
                    if not p_room:
                        user_name = forms.ask_for_string(
                            default=col,
                            prompt=u'找不到「{}」參數 (於房間 {})。\n請輸入正確的 Revit 參數名稱 (若是自訂參數)。\n留空代表強制略過此欄。'.format(col, room.Number)
                        )
                        if not user_name:
                            param_alias[col] = u''  # 記住略過
                            continue
                        revit_param_name = user_name
                    
                    param_alias[col] = revit_param_name

                if not revit_param_name:
                    continue

                # 執行寫入
                p = room.LookupParameter(revit_param_name)
                if p:
                    # 避免不必要的寫入 (例如原本值就一樣)
                    old_val = normalize(get_param_text(p))
                    if old_val != val:
                        try:
                            set_param_value(p, val)
                            room_updated = True
                        except Exception as e:
                            failures.append((u"房間: "+key_val, u"參數: "+revit_param_name, unicode(e)))
                else:
                    failures.append((u"房間: "+key_val, u"參數: "+revit_param_name, u"Revit 中無此參數"))

            if room_updated:
                updated_rooms += 1

        t.Commit()
    except Exception as e:
        t.RollBack()
        forms.alert(u'發生錯誤，已復原變更：{}'.format(e), exitscript=True); return

    # 6) 結果報告
    msg = [
        u'匯入與回寫完成！',
        u'- 成功更新房間數：{}'.format(updated_rooms),
        u'- Revit 找不到對應房間數：{}'.format(notfound_in_revit),
        u'- 鍵值空白而跳過行數：{}'.format(skipped_rows)
    ]
    if failures:
        msg.append(u'\n- 失敗/略過欄位：{} 次'.format(len(failures)))
        msg.append(u'  (例: {})'.format(failures[0]))
        
    forms.alert('\n'.join(msg), title=u'CSV 回寫結果', warn_icon=False)

if __name__ == '__main__':
    main()
