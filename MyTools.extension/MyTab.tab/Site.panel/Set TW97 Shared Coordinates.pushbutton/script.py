# -*- coding: utf-8 -*-
"""
Set TW97 Shared Coordinates
============================
透過 Revit ProjectLocation API 設定專案的 Shared Coordinates（共享座標），
以台灣 TWD97 二度分帶座標系統（X=Easting, Y=Northing）為輸入來源。

工作邏輯：
  1. 以 Internal Origin（XYZ.Zero）作為對應點
  2. 將公尺單位的座標轉換為 Revit 內部單位 feet
  3. 角度以 DMS（度/分/秒）+ East/West 方向輸入，與 Revit Site 對話框格式一致
  4. 呼叫 ProjectLocation.SetProjectPosition(XYZ.Zero, ProjectPosition(...))

注意：本工具僅修改座標設定，不移動任何模型幾何。

角度對應規則（與 Revit Site 對話框一致）：
  Revit UI「Angle from Project North to True North」
    West → angle_rad = -abs(angle_rad)   ← Project North 在 True North 東側
    East → angle_rad = +abs(angle_rad)   ← Project North 在 True North 西側

  ★ 若現場測試後發現方向相反，只需在 apply_project_position() 中
    交換 East/West 的正負號即可。

Author  : CP
Version : 1.1
Revit   : 2024
pyRevit : 4.8+
"""

import math
import clr

# ── Revit API ──────────────────────────────────────────────────────────────
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    Transaction,
    ProjectPosition,
    XYZ,
)

# ── pyRevit ────────────────────────────────────────────────────────────────
from pyrevit import script

# ── Windows Forms（多欄位輸入用）──────────────────────────────────────────
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System.Windows.Forms as WinForms
import System.Drawing as Drawing

# ── Revit 文件 ─────────────────────────────────────────────────────────────
doc = __revit__.ActiveUIDocument.Document


# ═══════════════════════════════════════════════════════════════════════════
#  常數
# ═══════════════════════════════════════════════════════════════════════════

METERS_TO_FEET = 3.28083989501312   # 1 公尺 = 3.28083... feet


# ═══════════════════════════════════════════════════════════════════════════
#  輔助函式
# ═══════════════════════════════════════════════════════════════════════════

def m_to_ft(value_m):
    """將公尺轉換為 Revit 內部單位（feet）。"""
    return value_m * METERS_TO_FEET


def dms_to_decimal(deg, minutes, sec):
    """將度分秒（DMS）轉換為十進位角度（decimal degrees）。"""
    return abs(deg) + abs(minutes) / 60.0 + abs(sec) / 3600.0


def parse_float(text, default=0.0, field_name=u"欄位"):
    """
    將字串解析為浮點數。
    - 空白時回傳 default
    - 非數字時拋出 ValueError（含欄位名稱）
    """
    text = text.strip()
    if text == "":
        return default
    try:
        return float(text)
    except ValueError:
        raise ValueError(u"「{}」輸入不合法，請輸入數字：{}".format(field_name, text))


def parse_int(text, default=0, field_name=u"欄位"):
    """
    將字串解析為整數（0–59 範圍）。
    - 空白時回傳 default
    - 非整數時拋出 ValueError
    """
    text = text.strip()
    if text == "":
        return default
    try:
        val = int(text)
    except ValueError:
        raise ValueError(u"「{}」請輸入整數（0–59）：{}".format(field_name, text))
    if val < 0 or val > 59:
        raise ValueError(u"「{}」必須在 0–59 之間，目前輸入：{}".format(field_name, val))
    return val


# ═══════════════════════════════════════════════════════════════════════════
#  輸入視窗（Windows Forms）
# ═══════════════════════════════════════════════════════════════════════════

# 版面常數
_PAD        = 20          # 左邊距
_LBL_W      = 180         # 標籤欄寬
_FIELD_X    = _PAD + _LBL_W  # 輸入欄起始 X
_FORM_W     = 520         # 視窗寬度
_ROW_H      = 34          # 列高
_FONT       = Drawing.Font("Segoe UI", 9)


def _label(text, x, y, w=_LBL_W, h=22):
    """建立並回傳一個 Label 控件。"""
    lbl = WinForms.Label()
    lbl.Text = text
    lbl.Location = Drawing.Point(x, y + 3)
    lbl.Width = w
    lbl.Height = h
    lbl.Font = _FONT
    return lbl


def _textbox(x, y, w=100, text=""):
    """建立並回傳一個 TextBox 控件。"""
    tb = WinForms.TextBox()
    tb.Location = Drawing.Point(x, y)
    tb.Width = w
    tb.Text = text
    tb.Font = _FONT
    return tb


class TW97InputForm(WinForms.Form):
    """
    多欄位輸入表單，角度格式與 Revit Location and Site 對話框一致：
      度（°）  分（'）  秒（"）  方向（East / West）
    """

    def __init__(self):
        self.Text = "Set TW97 Shared Coordinates"
        self.Width = _FORM_W
        self.AutoSize = True
        self.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.Padding = WinForms.Padding(0, 0, 0, 12)

        # 表單顯示後重設所有 TextBox 游標到最左側，避免文字被右切
        self.Shown += self._on_shown

        self._controls = []  # 暫存，最後一次 Add
        y = 16

        # ── Easting ──────────────────────────────────────────────────────
        y = self._add_coord_row(y,
            label=u"TW97 Easting / X (m)",
            attr="txt_easting",
            default="328266.280")

        # ── Northing ─────────────────────────────────────────────────────
        y = self._add_coord_row(y,
            label=u"TW97 Northing / Y (m)",
            attr="txt_northing",
            default="2747439.616")

        # ── Elevation ────────────────────────────────────────────────────
        y = self._add_coord_row(y,
            label=u"Elevation / Z (m)",
            attr="txt_elevation",
            default="0",
            hint=u"（空白預設 0）")

        # ── 分隔線 ────────────────────────────────────────────────────────
        sep = WinForms.Label()
        sep.BorderStyle = WinForms.BorderStyle.Fixed3D
        sep.Location = Drawing.Point(_PAD, y + 4)
        sep.Width = _FORM_W - _PAD * 2 - 16
        sep.Height = 2
        self._controls.append(sep)
        y += 18

        # ── Angle from Project North to True North（DMS + East/West）────
        lbl_section = _label(
            u"Angle from Project North to True North :",
            _PAD, y, w=_FORM_W - _PAD * 2 - 16)
        lbl_section.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        self._controls.append(lbl_section)
        y += _ROW_H - 4

        # DMS 欄位
        x = _FIELD_X

        # Degrees（最多 3 位，如 180）
        self.txt_deg = _textbox(x, y, w=58, text="0")
        self._controls.append(self.txt_deg)
        x += 62
        self._controls.append(_label(u"°", x - 8, y, w=16))

        # Minutes（0–59）
        self.txt_min = _textbox(x, y, w=50, text="0")
        self._controls.append(self.txt_min)
        x += 54
        self._controls.append(_label(u"'", x - 8, y, w=16))

        # Seconds（0–59.99）
        self.txt_sec = _textbox(x, y, w=62, text="0")
        self._controls.append(self.txt_sec)
        x += 66
        self._controls.append(_label(u'"', x - 8, y, w=16))

        # Direction ComboBox（East / West，與 Revit UI 一致）
        self.cmb_dir = WinForms.ComboBox()
        self.cmb_dir.Location = Drawing.Point(x, y)
        self.cmb_dir.Width = 80
        self.cmb_dir.Font = _FONT
        self.cmb_dir.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.cmb_dir.Items.Add("West")
        self.cmb_dir.Items.Add("East")
        self.cmb_dir.SelectedIndex = 0   # 預設 West（台灣多數專案）
        self._controls.append(self.cmb_dir)
        y += _ROW_H + 4

        # ── 提示文字 ──────────────────────────────────────────────────────
        lbl_hint = WinForms.Label()
        lbl_hint.Text = (
            u"※ 格式與 Revit「Location and Site」對話框相同。\n"
            u"※ 以 Internal Origin 為對應點，不移動模型幾何。\n"
            u"※ 完成後請以 Spot Coordinate 或 Survey Point 確認結果。"
        )
        lbl_hint.Location = Drawing.Point(_PAD, y)
        lbl_hint.Width = _FORM_W - _PAD * 2 - 16
        lbl_hint.Height = 52
        lbl_hint.Font = Drawing.Font("Segoe UI", 8)
        lbl_hint.ForeColor = Drawing.Color.Gray
        self._controls.append(lbl_hint)
        y += 58

        # ── OK / Cancel 按鈕 ──────────────────────────────────────────────
        btn_ok = WinForms.Button()
        btn_ok.Text = u"確定（Apply）"
        btn_ok.Location = Drawing.Point(_PAD, y)
        btn_ok.Width = 120
        btn_ok.Font = _FONT
        btn_ok.DialogResult = WinForms.DialogResult.OK
        self._controls.append(btn_ok)
        self.AcceptButton = btn_ok

        btn_cancel = WinForms.Button()
        btn_cancel.Text = u"取消"
        btn_cancel.Location = Drawing.Point(_PAD + 132, y)
        btn_cancel.Width = 80
        btn_cancel.Font = _FONT
        btn_cancel.DialogResult = WinForms.DialogResult.Cancel
        self._controls.append(btn_cancel)
        self.CancelButton = btn_cancel
        y += 38

        # 設定視窗高度（AutoSize GrowOnly 有時不可靠，手動補一下）
        self.ClientSize = Drawing.Size(_FORM_W - 16, y)

        # 一次加入全部控件
        for c in self._controls:
            self.Controls.Add(c)

    # ── Shown 事件：重設所有 TextBox 游標到最左 ──────────────────────────
    def _on_shown(self, sender, e):
        """表單顯示後把每個 TextBox 的游標移到最左側，避免預設值被右截。"""
        for tb in [self.txt_easting, self.txt_northing, self.txt_elevation,
                   self.txt_deg, self.txt_min, self.txt_sec]:
            tb.SelectionStart = 0
            tb.SelectionLength = 0

    # ── 座標列輔助 ────────────────────────────────────────────────────────
    def _add_coord_row(self, y, label, attr, default="", hint=""):
        """加入一列「標籤 + TextBox（+ 小提示）」，回傳下一列的 y。"""
        lbl = _label(label + (u"  " + hint if hint else ""), _PAD, y,
                     w=_LBL_W + (80 if hint else 0))
        self._controls.append(lbl)

        tb = _textbox(_FIELD_X + (80 if hint else 0), y, w=160, text=default)
        setattr(self, attr, tb)
        self._controls.append(tb)
        return y + _ROW_H

    # ── 讀取結果 ──────────────────────────────────────────────────────────
    @property
    def direction_west(self):
        """True → West；False → East"""
        return self.cmb_dir.SelectedItem == "West"


# ═══════════════════════════════════════════════════════════════════════════
#  輸入驗證
# ═══════════════════════════════════════════════════════════════════════════

def validate_inputs(form):
    """
    從表單讀取並驗證所有使用者輸入。
    回傳 (easting_m, northing_m, elevation_m, angle_decimal_deg, direction_west)
    若驗證失敗則拋出 ValueError。
    """
    # 座標欄位
    if form.txt_easting.Text.strip() == "":
        raise ValueError(u"TW97 Easting 不可為空白，請輸入數字。")
    if form.txt_northing.Text.strip() == "":
        raise ValueError(u"TW97 Northing 不可為空白，請輸入數字。")

    easting_m   = parse_float(form.txt_easting.Text,   field_name=u"TW97 Easting")
    northing_m  = parse_float(form.txt_northing.Text,  field_name=u"TW97 Northing")
    elevation_m = parse_float(form.txt_elevation.Text, default=0.0, field_name=u"Elevation")

    # 角度 DMS
    deg_val = parse_float(form.txt_deg.Text, default=0.0, field_name=u"角度（度）")
    min_val = parse_int(form.txt_min.Text,   default=0,   field_name=u"角度（分）")
    sec_val = parse_float(form.txt_sec.Text, default=0.0, field_name=u"角度（秒）")

    if deg_val < 0:
        raise ValueError(u"角度（度）不可為負數，方向請用 East / West 選擇。")
    if sec_val < 0 or sec_val >= 60:
        raise ValueError(u"角度（秒）必須在 0–59.99 之間。")

    angle_decimal = dms_to_decimal(deg_val, min_val, sec_val)
    direction_west = form.direction_west

    return easting_m, northing_m, elevation_m, angle_decimal, direction_west


# ═══════════════════════════════════════════════════════════════════════════
#  座標寫入
# ═══════════════════════════════════════════════════════════════════════════

def apply_project_position(doc, easting_m, northing_m, elevation_m,
                           angle_decimal_deg, direction_west):
    """
    透過 ProjectLocation.SetProjectPosition() 設定 Shared Coordinates。

    角度符號規則（對應 Revit「Location and Site」對話框）：
    ─────────────────────────────────────────────────────────
    Revit UI：「Angle from Project North to True North」
      West → Project North 在 True North 的東側
             → 從 True North 到 Project North 是順時針（正東方向）
             → ProjectPosition.Angle 取負值
      East → Project North 在 True North 的西側
             → 從 True North 到 Project North 是逆時針（正西方向）
             → ProjectPosition.Angle 取正值

    ★ 若現場測試後發現方向相反，只需交換下方兩行正負號即可。
    """
    easting_ft   = m_to_ft(easting_m)
    northing_ft  = m_to_ft(northing_m)
    elevation_ft = m_to_ft(elevation_m)

    angle_rad_abs = math.radians(abs(angle_decimal_deg))

    # ── 角度方向符號（★ 測試後可於此調整）────────────────────────────────
    if direction_west:
        angle_rad = -angle_rad_abs   # West → 負值
    else:
        angle_rad = +angle_rad_abs   # East → 正值

    project_location = doc.ActiveProjectLocation

    with Transaction(doc, "Set TW97 Shared Coordinates") as t:
        t.Start()
        # 以 Internal Origin（XYZ.Zero）為對應點，不移動模型幾何
        pos = ProjectPosition(easting_ft, northing_ft, elevation_ft, angle_rad)
        project_location.SetProjectPosition(XYZ.Zero, pos)
        t.Commit()

    return project_location.Name


# ═══════════════════════════════════════════════════════════════════════════
#  完成確認訊息
# ═══════════════════════════════════════════════════════════════════════════

def decimal_to_dms_str(decimal_deg):
    """將十進位角度轉回 D° M' S" 字串，方便確認訊息顯示。"""
    total_sec = round(abs(decimal_deg) * 3600, 2)
    d = int(total_sec // 3600)
    m = int((total_sec % 3600) // 60)
    s = total_sec % 60
    return u'{}° {}\' {:.2f}"'.format(d, m, s)


def show_result(easting_m, northing_m, elevation_m, angle_decimal_deg,
                direction_west, location_name):
    """跳出執行結果確認視窗。"""
    dir_str = "West" if direction_west else "East"
    dms_str = decimal_to_dms_str(angle_decimal_deg)
    msg = (
        u"✔ 座標設定完成\n\n"
        u"  Easting              : {:.3f} m\n"
        u"  Northing             : {:.3f} m\n"
        u"  Elevation            : {:.3f} m\n"
        u"  Angle (Project→True) : {} {}\n"
        u"  ProjectLocation      : {}\n\n"
        u"⚠ 請用 Spot Coordinate 或 Survey Point 顯示值再次確認。"
    ).format(
        easting_m, northing_m, elevation_m,
        dms_str, dir_str, location_name
    )
    WinForms.MessageBox.Show(
        msg,
        "Set TW97 Shared Coordinates",
        WinForms.MessageBoxButtons.OK,
        WinForms.MessageBoxIcon.Information
    )


# ═══════════════════════════════════════════════════════════════════════════
#  主程式
# ═══════════════════════════════════════════════════════════════════════════

def main():
    form = TW97InputForm()
    result = form.ShowDialog()

    if result != WinForms.DialogResult.OK:
        script.exit()
        return

    try:
        easting_m, northing_m, elevation_m, angle_dec, direction_west = validate_inputs(form)
    except ValueError as ex:
        WinForms.MessageBox.Show(
            str(ex),
            u"輸入錯誤",
            WinForms.MessageBoxButtons.OK,
            WinForms.MessageBoxIcon.Error
        )
        script.exit()
        return

    try:
        location_name = apply_project_position(
            doc, easting_m, northing_m, elevation_m, angle_dec, direction_west
        )
    except Exception as ex:
        WinForms.MessageBox.Show(
            u"寫入 Revit 時發生錯誤：\n{}".format(str(ex)),
            u"執行錯誤",
            WinForms.MessageBoxButtons.OK,
            WinForms.MessageBoxIcon.Error
        )
        script.exit()
        return

    show_result(easting_m, northing_m, elevation_m, angle_dec,
                direction_west, location_name)


if __name__ == "__main__":
    main()
