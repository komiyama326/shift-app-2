# pdf_exporter.py

import sys
import os
import calendar
import datetime
from dateutil.relativedelta import relativedelta
import holidays

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstallerが作成した一時フォルダのパスを取得
        base_path = sys._MEIPASS
    except Exception:
        # 開発環境（PyInstallerで実行されていない場合）のパスを取得
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- 定数定義 ---
FONT_NAME = "IPAexGothic"
FONT_BOLD_NAME = "IPAexGothic-Bold" # 太字用の別名を定義
FONT_FILE = "ipaexg.ttf"

try:
    # resource_path関数でフォントファイルの絶対パスを取得
    font_path = resource_path(FONT_FILE) 
    
    # 取得したパスを使ってフォントを登録
    pdfmetrics.registerFont(TTFont(FONT_NAME, font_path))
    # 'b'タグ（太字）用に同じフォントファイルを太字用の別名で登録
    pdfmetrics.registerFont(TTFont(FONT_BOLD_NAME, font_path))
    # フォントファミリーとして登録し、bold=で太字フォントの別名を指定
    pdfmetrics.registerFontFamily(FONT_NAME, normal=FONT_NAME, bold=FONT_BOLD_NAME)

except Exception as e:
    print(f"フォントの読み込みに失敗しました: {e}")
    print("PDF出力機能は正常に動作しない可能性があります。")
    # フォールバックフォントも念のため残しておく
    FONT_NAME = "HeiseiKakuGo-W5"

# --- メイン関数 ---
def export_to_pdf(filepath, year, month, title, schedule_data, staff_manager, format_type='grid', prev_month_schedule=None):
    # ... (変更なし)
    try:
        if format_type == 'grid':
            c = canvas.Canvas(filepath, pagesize=landscape(A4))
            width, height = landscape(A4)
            _generate_grid_format(c, width, height, year, month, title, schedule_data, staff_manager, prev_month_schedule)
        elif format_type == 'list':
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4
            _generate_list_format(c, width, height, year, month, title, schedule_data, staff_manager)
        else:
            raise ValueError("Unsupported format_type. Must be 'grid' or 'list'.")
        
        c.showPage()
        c.save()
        return True, None
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, str(e)


# ★★★★★ グリッド形式PDF生成（AttributeError修正版） ★★★★★
def _generate_grid_format(c, width, height, year, month, title, schedule_data, staff_manager, prev_month_schedule=None):
    margin_x = 15 * mm
    margin_y = 10 * mm

    # --- 色定義 ---
    sat_header_fill = colors.HexColor('#4F81BD')
    sun_header_fill = colors.HexColor('#C0504D')
    weekday_header_fill = colors.HexColor('#EAF1DD')
    multi_staff_fill = colors.HexColor('#FFFFE0')
    font_white = colors.white
    font_black = colors.black
    font_grey = colors.HexColor('#A9A9A9')
    
    # --- カレンダー描画エリア設定 ---
    cal_area_x = margin_x
    cal_area_width = width - 2 * margin_x
    col_width = cal_area_width / 7
    
    top_header_height = 12 * mm
    weekday_header_height = 9 * mm
    cal = calendar.monthcalendar(year, month)
    num_weeks = len(cal)
    body_height = height - (margin_y * 2) - top_header_height - weekday_header_height
    week_height = body_height / num_weeks
    cal_area_y = margin_y
    
    # --- ヘッダー描画 ---
    header_base_y = cal_area_y + body_height + weekday_header_height
    # ★★★★★ ParagraphStyleで太字フォント名を明示的に指定 ★★★★★
    style_header_main = ParagraphStyle(name='HeaderMain', fontName=FONT_BOLD_NAME, fontSize=20, alignment=TA_CENTER)
    p_title = Paragraph(title, style_header_main)
    p_month = Paragraph(f"{month}月", style_header_main)
    title_x = cal_area_x + col_width
    title_w = cal_area_width - col_width
    
    p_title.wrapOn(c, title_w, top_header_height)
    p_title.drawOn(c, title_x, header_base_y + (top_header_height - p_title.height) / 2)
    p_month.wrapOn(c, col_width, top_header_height)
    p_month.drawOn(c, cal_area_x, header_base_y + (top_header_height - p_month.height) / 2)
    
    # --- 曜日ヘッダー描画 ---
    weekdays_jp = ("月", "火", "水", "木", "金", "土", "日")
    header_y = cal_area_y + body_height
    for i, day_name in enumerate(weekdays_jp):
        x = cal_area_x + i * col_width
        if i == 5: fill_color, text_color = sat_header_fill, font_white
        elif i == 6: fill_color, text_color = sun_header_fill, font_white
        else: fill_color, text_color = weekday_header_fill, font_black
        c.setFillColor(fill_color)
        c.rect(x, header_y, col_width, weekday_header_height, fill=1, stroke=0)
        style = ParagraphStyle(name='Weekday', fontName=FONT_BOLD_NAME, fontSize=14, textColor=text_color, alignment=TA_CENTER)
        p = Paragraph(day_name, style)
        p.wrapOn(c, col_width, weekday_header_height)
        p.drawOn(c, x, header_y + (weekday_header_height - p.height) / 2)

    # --- カレンダー本体描画 ---
    jp_holidays = holidays.JP(years=year)
    date_ratio, staff_ratio, memo_ratio = 0.30, 0.45, 0.25

    for w_idx, week in enumerate(cal):
        for d_idx, day in enumerate(week):
            # ...(前月日付計算などは変更なし)...
            cell_x = cal_area_x + d_idx * col_width
            cell_y = cal_area_y + (num_weeks - w_idx - 1) * week_height
            
            staff_list = []
            is_current_month = (day != 0)
            if is_current_month:
                target_date = datetime.date(year, month, day)
                staff_list = schedule_data.get(target_date, [])
                display_day = str(day)
                day_text_color = font_black
            else:
                first_day_of_month = datetime.date(year, month, 1).weekday()
                if w_idx == 0 and d_idx < first_day_of_month and prev_month_schedule:
                    day_offset = first_day_of_month - d_idx
                    target_date = datetime.date(year, month, 1) - datetime.timedelta(days=day_offset)
                    staff_list = prev_month_schedule.get(target_date, [])
                    display_day = str(target_date.day)
                    day_text_color = font_grey
                else: continue

            memo_area_y = cell_y
            staff_area_y = memo_area_y + week_height * memo_ratio
            date_area_y = staff_area_y + week_height * staff_ratio
            date_area_height = week_height * date_ratio
            staff_area_height = week_height * staff_ratio
            
            date_bg_color = colors.transparent
            if is_current_month:
                is_holiday = target_date in jp_holidays
                is_saturday = target_date.weekday() == 5
                is_sunday = target_date.weekday() == 6
                if is_sunday or is_holiday: date_bg_color, day_text_color = sun_header_fill, font_white
                elif is_saturday: date_bg_color, day_text_color = sat_header_fill, font_white
            c.setFillColor(date_bg_color)
            c.rect(cell_x, date_area_y, col_width, date_area_height, fill=1, stroke=0)
            style_date = ParagraphStyle(name='Date', fontName=FONT_BOLD_NAME, fontSize=16, textColor=day_text_color, alignment=TA_CENTER)
            p_date = Paragraph(display_day, style_date)
            p_date.wrapOn(c, col_width, date_area_height)
            p_date.drawOn(c, cell_x, date_area_y + (date_area_height - p_date.height) / 2)

            staff_bg_color = colors.transparent
            staff_text_color = font_grey if not is_current_month else font_black
            if staff_list:
                if is_current_month:
                    if len(staff_list) == 1: staff_bg_color = colors.HexColor(staff_list[0].color_code)
                    else: staff_bg_color = multi_staff_fill
            c.setFillColor(staff_bg_color)
            c.rect(cell_x, staff_area_y, col_width, staff_area_height, fill=1, stroke=0)
            if staff_list:
                style_staff = ParagraphStyle(name='StaffStyle', fontName=FONT_BOLD_NAME, fontSize=20, leading=22, alignment=TA_CENTER, textColor=staff_text_color)
                text = "<br/>".join(staff_list) if isinstance(staff_list[0], str) else "<br/>".join([s.name for s in staff_list])
                p_staff = Paragraph(text, style_staff)
                p_w, p_h = p_staff.wrapOn(c, col_width - 2 * mm, staff_area_height)
                p_staff.drawOn(c, cell_x + 1 * mm, staff_area_y + (staff_area_height - p_h) / 2)

    # --- 罫線描画 ---
    c.setStrokeColor(font_black)
    c.setLineWidth(0.5)
    for i in range(1, 8):
        x = cal_area_x + i * col_width
        if i < 7:
            # トップヘッダーの縦線
            if i == 1:
                c.line(x, header_base_y, x, header_base_y + top_header_height)
            # 曜日ヘッダーの縦線
            c.line(x, header_y, x, header_y + weekday_header_height)
            # 本体の縦線
            c.line(x, cal_area_y, x, header_y)

    c.setLineWidth(0.2)
    for i in range(num_weeks):
        y = cal_area_y + i * week_height
        c.line(cal_area_x, y + week_height * memo_ratio, cal_area_x + cal_area_width, y + week_height * memo_ratio)
        c.line(cal_area_x, y + week_height * (memo_ratio + staff_ratio), cal_area_x + cal_area_width, y + week_height * (memo_ratio + staff_ratio))

    c.setLineWidth(1.5)
    c.rect(cal_area_x, cal_area_y, cal_area_width, body_height + weekday_header_height)
    c.rect(cal_area_x, header_base_y, cal_area_width, top_header_height)
    c.line(cal_area_x, header_y, cal_area_x + cal_area_width, header_y)
    for i in range(1, num_weeks):
        y = cal_area_y + i * week_height
        c.line(cal_area_x, y, cal_area_x + cal_area_width, y)



# --- タイムラインリスト形式PDFの生成 ---
def _generate_list_format(c, width, height, year, month, title, schedule_data, staff_manager):
    # ... (この関数は変更なし) ...
    margin = 20 * mm
    content_width = width - 2 * margin
    current_y = height - margin

    c.setFont(FONT_NAME, 18)
    c.drawCentredString(width / 2, current_y, title)
    current_y -= 8 * mm
    c.setFont(FONT_NAME, 14)
    c.drawCentredString(width / 2, current_y, f"{year}年 {month}月")
    current_y -= 12 * mm

    all_staff = sorted(staff_manager.get_all_staff(), key=lambda s: s.name)
    c.setFont(FONT_NAME, 11)
    c.drawString(margin, current_y, "凡例")
    current_y -= 6 * mm
    
    legend_x = margin
    c.setFont(FONT_NAME, 9)
    for staff in all_staff:
        staff_color = colors.HexColor(staff.color_code)
        if legend_x + 40 * mm > width - margin:
            legend_x = margin
            current_y -= 5 * mm
        c.setFillColor(staff_color)
        c.rect(legend_x, current_y - 0.5 * mm, 3*mm, 3*mm, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.drawString(legend_x + 5 * mm, current_y, staff.name)
        legend_x += 35 * mm
    current_y -= 10 * mm

    c.setStrokeColor(colors.lightgrey)
    c.line(margin, current_y, width - margin, current_y)
    current_y -= 5 * mm
    c.setFont(FONT_NAME, 9)
    c.setFillColor(colors.grey)
    c.drawString(margin, current_y, "DATE & DAY")
    c.drawString(margin + 30 * mm, current_y, "STAFF")
    c.setFillColor(colors.black)
    current_y -= 2 * mm
    
    jp_holidays = holidays.JP(years=year)
    _, num_days = calendar.monthrange(year, month)

    for day_num in range(1, num_days + 1):
        if current_y < margin + 10 * mm:
            c.showPage()
            current_y = height - margin

        c.setStrokeColor(colors.lightgrey)
        c.line(margin, current_y, width - margin, current_y)
        current_y -= 7 * mm

        date = datetime.date(year, month, day_num)
        staff_list = schedule_data.get(date, [])
        weekday_jp = ("月", "火", "水", "木", "金", "土", "日")[date.weekday()]

        is_saturday = date.weekday() == 5
        is_sunday = date.weekday() == 6
        is_holiday = date in jp_holidays
        day_color = colors.black
        if is_sunday or is_holiday: day_color = colors.crimson
        elif is_saturday: day_color = colors.royalblue

        c.setFont(FONT_NAME, 12)
        c.setFillColor(day_color)
        c.drawString(margin, current_y, f"{day_num:02d}")
        c.setFont(FONT_NAME, 10)
        c.drawString(margin + 8 * mm, current_y, f"({weekday_jp})")
        c.setFillColor(colors.black)

        staff_x = margin + 30 * mm
        c.setFont(FONT_NAME, 10)
        if not staff_list:
            c.setFillColor(colors.grey)
            c.drawString(staff_x, current_y, "(担当なし)")
            c.setFillColor(colors.black)
        else:
            for staff in staff_list:
                staff_color = colors.HexColor(staff.color_code)
                if staff_x + 35 * mm > width - margin:
                    staff_x = margin + 30 * mm
                    current_y -= 6 * mm
                c.setFillColor(staff_color)
                c.rect(staff_x, current_y - 0.5 * mm, 3*mm, 3*mm, fill=1, stroke=0)
                c.setFillColor(colors.black)
                c.drawString(staff_x + 5*mm, current_y, staff.name)
                staff_x += 35 * mm
        
        current_y -= 3 * mm