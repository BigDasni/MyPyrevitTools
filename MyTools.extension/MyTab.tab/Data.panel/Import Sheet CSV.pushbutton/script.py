# -*- coding: utf-8 -*-
# 這是0814的版本

# -*- coding: utf-8 -*-
from __future__ import print_function
from pyrevit import forms, revit, DB
import csv

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

# ---------- 讀 CSV（IronPython 友善：bytes 讀入後逐格 decode） ----------
# ---------- 改良版讀 CSV ----------
import csv

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
        except Exception as e:
            last_err = e
            continue
    raise Exception('CSV 讀取失敗（嘗試 utf-8-sig/utf-8/cp950）。最後錯誤：{}'.format(last_err))


# ---------- Revit 取值/寫值 ----------
def get_param_text(param):
    if not param or not param.HasValue:
        return None
    st = param.StorageType
    try:
        if st == DB.StorageType.String:
            return param.AsString()
        elif st == DB.StorageType.Integer:
            return unicode(param.AsInteger())
        elif st == DB.StorageType.Double:
            return unicode(param.AsDouble())
        else:
            return param.AsValueString()
    except:
        return None

def set_param_value(param, val):
    st = param.StorageType
    if st == DB.StorageType.String:
        param.Set(val)
    elif st == DB.StorageType.Integer:
        param.Set(int(val))
    elif st == DB.StorageType.Double:
        param.Set(float(val))
    else:
        param.Set(val)

# 只收集圖紙（ViewSheet）
def build_sheet_index_by_param(param_name):
    index = {}
    for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet):
        if not s: 
            continue
        p = s.LookupParameter(param_name)
        val = normalize(get_param_text(p))
        if val:
            index.setdefault(val, s)
    return index

# 找到圖紙上的標題欄（Title Block）實例
def get_titleblocks_on_sheet(sheet):
    cat = DB.BuiltInCategory.OST_TitleBlocks
    fi = DB.FilteredElementCollector(doc, sheet.Id).OfCategory(cat).WhereElementIsNotElementType()
    # 有些版本需要用全域 collector 過濾 OwnerViewId：
    if fi.GetElementCount() == 0:
        all_tb = DB.FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
        fi = [tb for tb in all_tb if getattr(tb, 'OwnerViewId', None) == sheet.Id]
        return fi
    return list(fi)

# ---------- 主流程 ----------
def main():
    csv_path = forms.pick_file(file_ext='csv', title='選擇要回寫的 CSV（UTF-8/Big5 皆可）')
    if not csv_path:
        forms.alert('已取消。'); return

    try:
        headers, rows, used_enc = read_csv_any(csv_path)
    except Exception as e:
        forms.alert('讀取 CSV 失敗：{}'.format(e), exitscript=True); return

    if not headers or not rows:
        forms.alert('CSV 無資料或無表頭。', exitscript=True); return

    # 1) 對應鍵（CSV 欄）→ 建議選「圖號」
    key_col = forms.ask_for_one_item(headers, title='選擇 CSV 的對應鍵欄（建議：圖號）', default='圖號')
    if not key_col:
        forms.alert('未選擇對應鍵。', exitscript=True); return

    # 2) Revit 參數名（預設 Sheet Number）
    revit_key_param = forms.ask_for_string(
        default='Sheet Number',
        prompt=u'輸入 Revit 內用來比對的圖紙參數名（預設：Sheet Number）。'
    )
    if not revit_key_param:
        forms.alert('未提供 Revit 參數名。', exitscript=True); return

    # 3) 要回寫的欄位
    # 3) 要回寫的欄位
    candidates = [h for h in headers if h != key_col]
    title_msg = u'選擇要回寫的欄位（可多選）\n---\n目前檔案編碼：{}'.format(used_enc)
    target_cols = forms.SelectFromList.show(candidates, multiselect=True, title=title_msg)

    if not target_cols:
        forms.alert('未選擇回寫欄位。', exitscript=True); return

    # 4) 建立圖紙索引
    forms.alert(u'建立圖紙索引：依 [{}] 比對…'.format(revit_key_param), warn_icon=False)
    sheet_index = build_sheet_index_by_param(revit_key_param)

    # 5) 欄位映射快取（CSV欄名 -> Revit參數名），避免每列都詢問
    param_alias = {}

    updated = notfound = skipped = 0
    failures = []

    t = DB.Transaction(doc, u'CSV 回寫圖紙/標題欄')
    t.Start()
    try:
        for r in rows:
            key_val = normalize(r.get(key_col))
            if not key_val:
                skipped += 1
                continue

            sheet = sheet_index.get(key_val)
            if not sheet:
                notfound += 1
                continue

            # 找到此圖紙上的所有標題欄
            titleblocks = get_titleblocks_on_sheet(sheet)

            for col in target_cols:
                raw_val = r.get(col)
                if raw_val is None:
                    continue
                val = normalize(raw_val)

                # 取得對應的 Revit 參數名（快取）
                revit_param_name = param_alias.get(col)
                if not revit_param_name:
                    # 先假設同名
                    revit_param_name = col
                    # 如果圖紙與任一標題欄都沒有這參數，才詢問一次映射
                    p_sheet = sheet.LookupParameter(revit_param_name)
                    p_tb = None
                    if not p_sheet and titleblocks:
                        for tb in titleblocks:
                            p_tb = tb.LookupParameter(revit_param_name)
                            if p_tb: break
                    if not p_sheet and not p_tb:
                        user_name = forms.ask_for_string(
                            default=col,
                            prompt=u'圖紙 {} 找不到「{}」參數。\n請輸入要寫入的參數名（會記住，下次不再詢問；留空略過）'
                                   .format(sheet.SheetNumber, col)
                        )
                        if not user_name:
                            # 記成空，後續直接略過此欄
                            param_alias[col] = u''
                            continue
                        revit_param_name = user_name
                    # 記憶映射
                    param_alias[col] = revit_param_name

                if not revit_param_name:
                    continue

                # 先嘗試寫到「圖紙」參數
                p = sheet.LookupParameter(revit_param_name)
                if p:
                    try:
                        set_param_value(p, val); updated += 1
                        continue
                    except Exception as e:
                        failures.append((sheet.SheetNumber, revit_param_name, unicode(e)))
                        continue

                # 圖紙沒有，就嘗試寫到「標題欄」參數（取第一個有此參數的標題欄）
                wrote = False
                for tb in titleblocks:
                    ptb = tb.LookupParameter(revit_param_name)
                    if ptb:
                        try:
                            set_param_value(ptb, val); updated += 1
                            wrote = True
                            break
                        except Exception as e:
                            failures.append((sheet.SheetNumber, revit_param_name, unicode(e)))
                            wrote = True
                            break
                if not wrote:
                    failures.append((sheet.SheetNumber, revit_param_name, u'在圖紙與標題欄皆無此參數'))

        t.Commit()
    except Exception as e:
        t.RollBack()
        forms.alert('發生錯誤，已回復：{}'.format(e), exitscript=True); return

    msg = [
        u'CSV 回寫完成（使用編碼：{}）：'.format(used_enc),
        u'- 成功寫入筆數：{}'.format(updated),
        u'- 找不到圖紙：{}'.format(notfound),
        u'- 跳過（鍵值空白）：{}'.format(skipped)
    ]
    if failures:
        msg.append(u'- 失敗/略過說明：{}'.format(len(failures)))
        msg.append(u'  例：{}'.format(failures[0]))
    forms.alert('\n'.join(msg), warn_icon=False)

if __name__ == '__main__':
    main()
