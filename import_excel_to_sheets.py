#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script: import_excel_to_sheets.py
คำอธิบาย: สคริปต์นำเข้าไฟล์ Excel ISDP (56A0S0Q และ 56A0UPS)
          ส่งขึ้น Google Spreadsheet ผ่าน Google Apps Script Web App API
          พร้อมสั่งรีเฟรชคำนวณแดชบอร์ด SLA อัตโนมัติ
"""

import os
import sys
import glob
import json
import urllib.request
from datetime import datetime
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

WEB_APP_URL = "https://script.google.com/macros/s/AKfycbysOK_GAlsnJ12VOLUUm-0qDltWipjba_JKYc2gdzE9M50FaGQ5O-R8gPiqEQK0LopsQQ/exec"

def format_date(val):
    if not val:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s or s.lower() == "none":
        return None
    return s.split(" ")[0]

def parse_ais_excel(file_path):
    print(f"\n📂 กำลังอ่านไฟล์ AIS (56A0S0Q): {os.path.basename(file_path)}...")
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet_name = next((s for s in wb.sheetnames if "site rollout plan" in s.lower()), wb.sheetnames[0])
    ws = wb[sheet_name]

    records = []
    # Row 1-2: Headers, Row 3: Subtotals, Row 4+: Data
    for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if r_idx < 3:
            continue
        duid = str(row[0]).strip() if row[0] is not None else ""
        if not duid or duid.startswith("="):
            continue

        install_date = format_date(row[8])
        if not install_date:
            continue

        smart_qc_date = format_date(row[20])
        pat_date = format_date(row[30])
        owner = str(row[31] or row[9] or "-").strip()
        pat_remark = str(row[38] or "").strip()
        pat_status = str(row[39] or "").strip()

        records.append({
            "duid": duid,
            "installDate": install_date,
            "smartQcDate": smart_qc_date,
            "patDate": pat_date,
            "owner": owner,
            "patRemark": pat_remark,
            "patStatus": pat_status
        })

    wb.close()
    print(f"✅ อ่านสำเร็จ: พบงานติดตั้ง AIS ทั้งหมด {len(records)} ไซต์")
    return records

def parse_true_excel(file_path):
    print(f"\n📂 กำลังอ่านไฟล์ TRUE (56A0UPS): {os.path.basename(file_path)}...")
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet_name = next((s for s in wb.sheetnames if "site rollout plan" in s.lower()), wb.sheetnames[0])
    ws = wb[sheet_name]

    records = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if r_idx < 3:
            continue
        duid = str(row[0]).strip() if row[0] is not None else ""
        if not duid or duid.startswith("="):
            continue

        install_date = format_date(row[66]) # Verify Photo Actual End Date
        if not install_date:
            continue

        smart_qc_date = format_date(row[36]) # 08.1 SmartQC
        aor_date = format_date(row[40])       # AOR
        alarm_remark = str(row[44] or "").strip() # Integration alarm category
        pat_date = format_date(row[46])       # 14.1 A129 PAT
        owner = str(row[47] or row[37] or row[3] or "-").strip()

        records.append({
            "duid": duid,
            "installDate": install_date,
            "smartQcDate": smart_qc_date,
            "aorDate": aor_date,
            "alarmRemark": alarm_remark,
            "patDate": pat_date,
            "owner": owner
        })

    wb.close()
    print(f"✅ อ่านสำเร็จ: พบงานติดตั้ง TRUE ทั้งหมด {len(records)} ไซต์")
    return records

def send_to_google_sheet(operator, records, chunk_size=300):
    if not records:
        print(f"⚠️ ไม่มีข้อมูลสำหรับส่งขึ้น {operator}")
        return

    total = len(records)
    print(f"🚀 กำลังส่งข้อมูล {operator} ไปยัง Google Sheets ({total} แถว)...")

    # ส่งเป็นชุด (Batch) เพื่อป้องกัน Apps Script Timeout
    for i in range(0, total, chunk_size):
        chunk = records[i:i + chunk_size]
        payload = json.dumps({
            "action": "saveImportData",
            "operator": operator,
            "records": chunk
        }).encode("utf-8")

        req = urllib.request.Request(
            WEB_APP_URL,
            data=payload,
            headers={"Content-Type": "application/json;charset=utf-8"}
        )

        try:
            with urllib.request.urlopen(req, timeout=45) as resp:
                res_body = resp.read().decode("utf-8")
                print(f"  -> Batch {i+1} ถึง {min(i+chunk_size, total)}: สำเร็จ ({res_body[:80]})")
        except Exception as e:
            print(f"  ❌ เกิดข้อผิดพลาดใน Batch {i+1}-{min(i+chunk_size, total)}: {e}")

    print(f"🎉 ส่งข้อมูล {operator} ขึ้น Google Spreadsheet เรียบร้อยแล้ว!\n")

def main():
    print("=" * 60)
    print("  🚀 ISDP EXCEL TO GOOGLE SHEETS IMPORTER (56A0S0Q & 56A0UPS)")
    print("=" * 60)

    # ค้นหาไฟล์ AIS (56A0S0Q)
    ais_files = glob.glob("*56A0S0Q*.xlsx") + glob.glob("*56a0s0q*.xlsx")
    if ais_files:
        ais_records = parse_ais_excel(ais_files[0])
        send_to_google_sheet("AIS", ais_records)
    else:
        print("⚠️ ไม่พบไฟล์ Excel ของ AIS (56A0S0Q)")

    # ค้นหาไฟล์ TRUE (56A0UPS)
    true_files = glob.glob("*56A0UPS*.xlsx") + glob.glob("*56a0ups*.xlsx")
    if true_files:
        true_records = parse_true_excel(true_files[0])
        send_to_google_sheet("TRUE", true_records)
    else:
        print("⚠️ ไม่พบไฟล์ Excel ของ TRUE (56A0UPS)")

    print("=" * 60)
    print("✨ กระบวนการนำเข้าข้อมูลและอัปเดต Dashboard เสร็จสมบูรณ์แล้ว!")
    print("=" * 60)

if __name__ == "__main__":
    main()
