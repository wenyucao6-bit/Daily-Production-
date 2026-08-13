#!/usr/bin/env python
# coding: utf-8

# In[21]:


# In[20]:


import requests
import time
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import shutil
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


# In[12]:


API_TOKEN_CGNEE  = "f90ad38034dd320aca9e447e99241a76"
BASE_URL_CGNEE   = "https://cgnee.greenbyte.cloud/api/2.0"

EMAIL_PASSWORD   = "Yv.13906465361"   

TEMPLATE_PATH    = r"C:\Users\CAO WENYU\OneDrive - CGN Europe Energy\Pro Template - Copy.xlsx"   # ← 改成你的模板路径


# In[13]:


# 模板里风场名称 → Greenbyte 风场名称 映射
# 模板用下划线，Greenbyte 用空格/连字符
# ─────────────────────────────────────────────
WF_NAME_MAP = {
    # Alize
    "CECOM_I":                          "CECOM I",
    "CECOM_II":                         "CECOM II",
    "CEHOC":                            "CEHOC",
    "CELBC":                            "CELBC",
    "CETOU_II":                         "CETOU II",
    "CETOU_IV":                         "CETOU IV",
    "CEVIN_I":                          "CEVIN I",
    "CEVIN_II":                         "CEVIN II",
    # Fujin
    "Falaise":                          "Falaise",
    "Vent_de_Gavray_-_Gratot":          "Vent de Gavray - Gratot",
    "Vent_de_Gavray_-_Sourdeval":       "Vent de Gavray - Sourdeval",
    "Vent_de_Gavray_-_Gavray":          "Vent de Gavray - Gavray",
    "Magoar":                           "Magoar",
    "Vix":                              "Vix",
    "Levigny":                          "Levigny",
    "Charmont-Ventelec":                "Charmont-Ventelec",
    "Saulzet_1":                        "Saulzet 1",
    "Saulzet_2":                        "Saulzet 2",
    "Scaer":                            "Scaer",
    # Hermes
    "Haute-Somme":                      "Haute-Somme",
    "SSH":                              "SSH",
    # Olympia
    "Biesles":                          "Biesles",
    "Basse_Tierache_Sud":               "Basse Tierache Sud",
    "Paradis_du_Plessis":               "Paradis du Plessis",
    # Windfall
    "Ext._Champfleury_-_Champfleury_2": "Ext. Champfleury - Champfleury 2",
    "Ext._Champfleury_-_Viapres_1":     "Ext. Champfleury - Viapres 1",
    "Ext._Champfleury_-_Viapres_2":     "Ext. Champfleury - Viapres 2",
    "Mont_dArcis_-_Allibaudieres":      "Mont d'Arcis - Allibaudieres",
    "Mont_dArcis_-_Le_Chene":           "Mont d'Arcis - Le Chene",
    "Mont_dArcis_-_Orme_Boyard":        "Mont d'Arcis - Orme Boyard",
    "Mont_dArcis_-_Dosnon":             "Mont d'Arcis - Dosnon",
    "Mont_dArcis_-_Vignes_Hautes":      "Mont d'Arcis - Vignes Hautes",
    "Quittebeuf":                       "Quittebeuf",
    "Assac_-_Puech_dAl_Lun":            "Assac - Puech d'Al Lun",
    "Assac_-_Garrigade":                "Assac - Garrigade",
    "Quatre_Vents_-_Cote_Guillaume":    "Quatre Vents - Cote Guillaume",
    "Quatre_Vents_-_Les_Vignes":        "Quatre Vents - Les Vignes",
    "Quatre_Vents_-_Norvilliers":       "Quatre Vents - Norvilliers",
    "Quatre_Vents_-_Couveillons":       "Quatre Vents - Couveillons",
    # Esperance
    "Estinnes":                         "Estinnes",
    # Forum
    "Maanderbroek":                     "Maanderbroek",
    "Nieuw_Prinsenland":                "Nieuw Prinsenland",
    "Oud_Dintel":                       "Oud Dintel",
    "Treurenburg":                      "Treurenburg",
    "Suurhoffbrug":                     "Suurhoffbrug",
    # Douvan NI (Northern Ireland)
    "Carn_Hill":                        "Carn Hill",
    "Dunbeg":                           "Dunbeg",
    "Monnaboy":                         "Monnaboy",
    "Cloonty":                          "Cloonty",
    "Corby_Knowe":                      "Corby Knowe",
    "Inishative":                       "Inishative",
    "Cregganconroe":                    "Cregganconroe",
    # Clover
    "Green_Rigg":                       "Green Rigg",
    "Glass_Moor_2":                     "Glass Moor 2",
    "Rusholme":                         "Rusholme",
    # Brenig / Paulette
    "Brenig":                           "Brenig",
    # Douvan IR (Ireland)
    "Ballywater":                       "Ballywater",
    "Cloghboola":                       "Cloghboola",
    "Faughary":                         "Faughary",
    "Leabeg":                           "Leabeg",
    "Roosky":                           "Roosky",
    "Skrine":                           "Skrine",
    "Ballybay":                         "Ballybay",
    # PowerHouse (Sweden) 
    "Högkölen_(Odin)":                  "Högkölen",
    "Lehtirova_(Thor)":                 "Lehtirova",
    "Nylandsbergen_(Small_Zlatan)":     "Nylandsbergen",
    "Kraktorpet_(Big_Zlatan)":          "Kraktorpet",
    "Valhalla_(Amot-Lingbo__Tonsen)":   "Valhalla",
}


# In[14]:


# ─────────────────────────────────────────────
# Step 1: 获取所有设备信息 (deviceId + site title)
# ─────────────────────────────────────────────
def get_devices(base_url, token):
    """获取风机设备列表，返回 {wind_farm_name: [device_id, ...]}"""
    url = f"{base_url}/devices.json"
    headers = {"Breeze-ApiToken": token}
    params = {
        "deviceTypeIds": 1,   # 1 = Wind turbine
        "fields": ["site", "title", "deviceId"],
        "pageSize": 1000
    }
    for attempt in range(2):
        try:
            r = requests.get(url=url, headers=headers, params=params, timeout=60)
            data = r.json()
            break
        except Exception:
            if attempt == 0:
                time.sleep(30)
            else:
                raise
 
    df = pd.DataFrame(data)
    df["site_name"] = df["site"].apply(lambda x: x["title"])
    result = df.groupby("site_name")["deviceId"].apply(list).to_dict()
    return result


# In[15]:


# ─────────────────────────────────────────────
# Step 2: 获取某风场某时间段的日产量 (MWh/天)
# ─────────────────────────────────────────────
def get_daily_production(base_url, token, device_ids, date_start, date_end):
    """
    拉取 date_start ~ date_end 每天的 Energy Export 产量
    返回 {date_str: MWh}，date_str 格式 'YYYY-MM-DD'
    """
    # 先查信号ID
    url_sig = f"{base_url}/datasignals.json"
    headers = {"Breeze-ApiToken": token}
    for attempt in range(2):
        try:
            r = requests.get(url=url_sig, headers=headers,
                             params={"deviceIds": device_ids}, timeout=60)
            sig_df = pd.DataFrame(r.json())
            break
        except Exception:
            if attempt == 0:
                time.sleep(30)
            else:
                return {}
 
    energy_sig = sig_df[sig_df["title"] == "Energy Export"]["dataSignalId"].tolist()
    if not energy_sig:
        print(f"    ⚠️  No 'Energy Export' signal found for devices {device_ids[:3]}...")
        return {}
 
    # 拉数据
    url_data = f"{base_url}/data.json"
    params = {
        "deviceIds": device_ids,
        "dataSignalIds": energy_sig,
        "timestampStart": date_start.strftime("%Y-%m-%dT%H:%M:%S"),
        "timestampEnd":   date_end.strftime("%Y-%m-%dT%H:%M:%S"),
        "aggregate":      "site",
        "resolution":     "daily",
    }
    for attempt in range(2):
        try:
            r = requests.get(url=url_data, headers=headers, params=params, timeout=60)
            df = pd.DataFrame(r.json())
            break
        except Exception:
            if attempt == 0:
                time.sleep(30)
            else:
                return {}
 
    if df.empty or "data" not in df.columns:
        return {}
 
    raw = df["data"].iloc[0]  # {timestamp_str: value_kWh}
    result = {}
    for ts_str, val in raw.items():
        date_str = ts_str[:10]   # 取 YYYY-MM-DD
        result[date_str] = val   # 保留 kWh
    return result


# In[16]:


# ─────────────────────────────────────────────
# Step 3: 拉取所有风场前一周产量
# ─────────────────────────────────────────────
def fetch_all_production(week_start, week_end):
    """
    week_start: Monday 00:00
    week_end:   Sunday 23:59 (or Monday next week 00:00)
    返回 {template_wf_name: {date_str: kWh}}
    """
    print("正在获取设备列表...")
    devices_cgnee = get_devices(BASE_URL_CGNEE, API_TOKEN_CGNEE)
 
    all_production = {}
 
    for template_name, gb_name in WF_NAME_MAP.items():
        device_ids = devices_cgnee.get(gb_name)
        if not device_ids:
            # 尝试部分匹配
            matches = [k for k in devices_cgnee if gb_name.lower() in k.lower()]
            if matches:
                device_ids = devices_cgnee[matches[0]]
                print(f"  ⚠️  '{gb_name}' 模糊匹配到 '{matches[0]}'")
            else:
                print(f"  ❌  '{gb_name}' 在 API 中未找到，跳过")
                all_production[template_name] = {}
                continue
 
        print(f"  拉取: {gb_name} ({len(device_ids)} 台风机)...")
        daily = get_daily_production(BASE_URL_CGNEE, API_TOKEN_CGNEE, device_ids, week_start, week_end)
        all_production[template_name] = daily
 
    return all_production


# In[17]:

# ─────────────────────────────────────────────
# Step 4: 写入 Excel 模板
# ─────────────────────────────────────────────
def write_report(production_data, week_dates, output_path):
    """
    week_dates: list of datetime.date objects (7天)
    production_data: {template_wf_name: {date_str: MWh}}
    """
    shutil.copy(TEMPLATE_PATH, output_path)
    wb = load_workbook(output_path)
    ws = wb["Sheet1"]
 
    # 写日期表头 (第1行, 从D列=第4列开始)
    header_fill  = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_side    = Side(style="thin", color="CCCCCC")
    thin_border  = Border(left=thin_side, right=thin_side,
                          top=thin_side, bottom=thin_side)
 
    for i, d in enumerate(week_dates):
        col = 4 + i
        cell = ws.cell(row=1, column=col)
        cell.value = d.strftime("%Y-%m-%d")
        cell.font  = header_font
        cell.fill  = header_fill
        cell.alignment = center_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col)].width = 13
 
    # ── 写数据行 ──
    data_font    = Font(name="Arial", size=10)
    num_format   = '#,##0.00'
    alt_fill     = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
 
    for excel_row in range(2, ws.max_row + 1):
        template_name = ws.cell(row=excel_row, column=3).value
        if not template_name:
            continue
 
        daily    = production_data.get(template_name, {})
        row_fill = alt_fill if (excel_row % 2 == 0) else None
 
        for i, d in enumerate(week_dates):
            col      = 4 + i
            date_str = d.strftime("%Y-%m-%d")
            val      = daily.get(date_str)
 
            cell = ws.cell(row=excel_row, column=col)
            cell.number_format = num_format
            cell.font          = data_font
            cell.border        = thin_border
            cell.alignment     = Alignment(horizontal="right")
            if row_fill:
                cell.fill = row_fill
            cell.value = round(val, 2) if val is not None else 0
 
    # ── 按模板A列(Country)分组，汇总国家产量，写在风场明细下方 ──
    country_rows = {}       # {country: [excel_row, ...]}
    current_country = None
    last_data_row = 1
    for excel_row in range(2, ws.max_row + 1):
        if not ws.cell(row=excel_row, column=3).value:   # C列没有风场名就跳过
            continue
        a_val = ws.cell(row=excel_row, column=1).value    # A列有值代表新国家开始
        if a_val:
            current_country = a_val
        if current_country:
            country_rows.setdefault(current_country, []).append(excel_row)
        last_data_row = excel_row

    if country_rows:
        country_header_fill = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
        country_header_font = Font(name="Arial", bold=True, size=12)
        country_label_fill  = PatternFill("solid", start_color="D9D9D9", end_color="D9D9D9")
        country_label_font  = Font(name="Arial", bold=True, size=10)

        header_row = last_data_row + 2   # 空一行分隔

        # "Country" 表头单元格 (A:C 合并)
        ws.merge_cells(start_row=header_row, start_column=1, end_row=header_row, end_column=3)
        hdr = ws.cell(row=header_row, column=1, value="Country")
        hdr.font = country_header_font
        hdr.fill = country_header_fill
        hdr.alignment = center_align
        hdr.border = thin_border

        # 沿用第1行的日期表头样式
        for i, d in enumerate(week_dates):
            col = 4 + i
            cell = ws.cell(row=header_row, column=col)
            cell.value = d.strftime("%Y-%m-%d")
            cell.font  = header_font
            cell.fill  = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        # 各国家求和行
        for offset, (country, rows_list) in enumerate(country_rows.items(), start=1):
            r = header_row + offset
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
            lbl = ws.cell(row=r, column=1, value=country)
            lbl.font = country_label_font
            lbl.fill = country_label_fill
            lbl.alignment = center_align
            lbl.border = thin_border

            min_r, max_r = min(rows_list), max(rows_list)
            contiguous = (max_r - min_r + 1) == len(rows_list)

            for i, d in enumerate(week_dates):
                col = 4 + i
                col_letter = get_column_letter(col)
                cell = ws.cell(row=r, column=col)
                if contiguous:
                    cell.value = f"=SUM({col_letter}{min_r}:{col_letter}{max_r})"
                else:
                    refs = ",".join(f"{col_letter}{rr}" for rr in rows_list)
                    cell.value = f"=SUM({refs})"
                cell.number_format = num_format
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="right")

    # 固定前3列
    ws.freeze_panes = "D2"
 
    wb.save(output_path)
    print(f"\n✅ 报告已保存: {output_path}")
    
# # ─────────────────────────────────────────────
# # Step 4: 写入 Excel 模板
# # ─────────────────────────────────────────────
# def write_report(production_data, week_dates, output_path):
#     """
#     week_dates: list of datetime.date objects (7天)
#     production_data: {template_wf_name: {date_str: MWh}}
#     """
#     shutil.copy(TEMPLATE_PATH, output_path)
#     wb = load_workbook(output_path)
#     ws = wb["Sheet1"]
 
#     # 写日期表头 (第1行, 从D列=第4列开始)
#     header_fill  = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
#     header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#     center_align = Alignment(horizontal="center", vertical="center")
#     thin_side    = Side(style="thin", color="CCCCCC")
#     thin_border  = Border(left=thin_side, right=thin_side,
#                           top=thin_side, bottom=thin_side)
 
#     for i, d in enumerate(week_dates):
#         col = 4 + i
#         cell = ws.cell(row=1, column=col)
#         cell.value = d.strftime("%Y-%m-%d")
#         cell.font  = header_font
#         cell.fill  = header_fill
#         cell.alignment = center_align
#         cell.border = thin_border
#         ws.column_dimensions[get_column_letter(col)].width = 13
 
#     # ── 写数据行 ──
#     data_font    = Font(name="Arial", size=10)
#     num_format   = '#,##0.00'
#     alt_fill     = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
 
#     for excel_row in range(2, ws.max_row + 1):
#         template_name = ws.cell(row=excel_row, column=3).value
#         if not template_name:
#             continue
 
#         daily    = production_data.get(template_name, {})
#         row_fill = alt_fill if (excel_row % 2 == 0) else None
 
#         for i, d in enumerate(week_dates):
#             col      = 4 + i
#             date_str = d.strftime("%Y-%m-%d")
#             val      = daily.get(date_str)
 
#             cell = ws.cell(row=excel_row, column=col)
#             cell.number_format = num_format
#             cell.font          = data_font
#             cell.border        = thin_border
#             cell.alignment     = Alignment(horizontal="right")
#             if row_fill:
#                 cell.fill = row_fill
#             cell.value = round(val, 2) if val is not None else 0
 
#     # 固定前3列
#     ws.freeze_panes = "D2"
 
#     wb.save(output_path)
#     print(f"\n✅ 报告已保存: {output_path}")


# In[22]:


def send_email(output_file, week_dates):
    msg = MIMEMultipart()
    msg["From"]    = SMTP_USER
    msg["To"]      = EMAIL_TO
    msg["Subject"] = f"Daily Production Report {week_dates[0].strftime('%Y-%m-%d')} ~ {week_dates[-1].strftime('%Y-%m-%d')}"

    body = f"""Dear All,

Please find attached the daily production report for the period {week_dates[0].strftime('%Y-%m-%d')} to {week_dates[-1].strftime('%Y-%m-%d')}.

Best regards,
Wenyu
"""
    msg.attach(MIMEText(body, "plain"))

    with open(output_file, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(output_file))
    part["Content-Disposition"] = f'attachment; filename="{os.path.basename(output_file)}"'
    msg.attach(part)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SMTP_USER, EMAIL_TO.split(";"), msg.as_string())

    print("✅ 邮件发送成功")


# In[23]:


# ─────────────────────────────────────────────
# 主函数
# ─────────────────────────────────────────────
def main():
    """
    每次运行自动取昨天往前推7天（滚动窗口）
    例如今天是 2026-06-01，则报告范围是 2026-05-25 ~ 2026-05-31
    """
    today     = datetime.today().date()
    week_end_date   = today - timedelta(days=1)          # 昨天
    week_start_date = week_end_date - timedelta(days=6)  # 往前推7天
 
    week_dates = [week_start_date + timedelta(days=i) for i in range(7)]
    week_start = datetime.combine(week_dates[0],  datetime.min.time())
    week_end   = datetime.combine(week_dates[-1], datetime.min.time()) + timedelta(days=1)
 
    #print(f"📅 报告周期: {week_dates[0]} ~ {week_dates[-1]} (前7天滚动)")
    print(f"   运行日期: {today}\n")
 
    production_data = fetch_all_production(week_start, week_end)
 
    output_file = os.path.join(os.path.expanduser("~"), "Downloads",
                  f"Weekly_Production_{week_dates[0].strftime('%Y%m%d')}_{week_dates[-1].strftime('%Y%m%d')}.xlsx")
    write_report(production_data, week_dates, output_file)
 
    print("\n📊 各风场每日产量 (kWh):")
    header = f"{'Wind farm':<45}" + "".join(f"  {d.strftime('%m-%d'):>10}" for d in week_dates)
    print(header)
    print("-" * len(header))
    for template_name, daily in production_data.items():
        row = f"{template_name:<45}"
        for d in week_dates:
            val = daily.get(d.strftime("%Y-%m-%d"))
            row += f"  {val:>10.1f}" if val is not None else f"  {0:>10.1f}"
        print(row)
 
    send_email(output_file, week_dates)
 
if __name__ == "__main__":
    main()


# In[ ]:




