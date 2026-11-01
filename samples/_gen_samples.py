"""產生範本與假資料（全部虛構）。用法：uv run --with openpyxl python samples/_gen_samples.py"""
import datetime as dt
import os
import random

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

random.seed(7)
HERE = os.path.dirname(os.path.abspath(__file__))
BASE = dt.date(2026, 11, 1)
HEAD = PatternFill("solid", fgColor="EEF1F5")
TRAININGS = ["保全職前專業訓練", "保全在職訓練", "一般安全衛生教育訓練（新僱）", "一般勞工安全衛生在職教育訓練",
             "職業安全衛生業務主管在職訓練", "急救人員在職訓練", "防火管理人初訓", "防火管理人複訓",
             "公寓大廈管理服務人員回訓講習", "公寓大廈管理服務人員認可證"]
ROSTER_COLS = ["員工編號", "姓名", "公司", "社區", "職稱", "到職日", "離職日", "離職原因", "職務標籤"]
TRAIN_COLS = ["員工編號", "姓名", "訓練項目", "完成日", "時數", "證號", "證照到期日"]
APPROVAL_COLS = ["單號", "單據名稱", "申請人", "公司", "社區", "申請日期", "目前簽核人", "狀態", "完成日期", "備註"]


def header(ws, cols, widths=None):
    ws.append(cols)
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
        c.fill = HEAD
    for i, _ in enumerate(cols):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = (widths or {}).get(i, 14)


def rnd_date(a, b):
    return a + dt.timedelta(days=random.randint(0, (b - a).days))


# ── 員工名冊＋訓練紀錄（假資料）──
companies = [("甲保全", 150, ["測試社區A", "測試社區B", "測試社區C", "測試大樓D"]),
             ("乙公寓大廈管理", 90, ["測試社區E", "測試社區F", "測試大樓G"]),
             ("丙清潔", 60, ["測試社區H", "測試大樓G"])]
reasons = ["個人因素", "薪資", "家庭因素", "轉職", "健康因素", "工時太長", "試用期未通過", "通勤距離"]
people, trains = [], []
n = 0
for comp, count, sites in companies:
    for _ in range(count):
        n += 1
        site = random.choice(sites)
        title = {"甲保全": random.choice(["保全員"] * 6 + ["保全組長", "機動保全"]),
                 "乙公寓大廈管理": random.choice(["總幹事", "秘書", "機電技師", "行政人員"]),
                 "丙清潔": random.choice(["清潔員"] * 5 + ["清潔組長"])}[comp]
        hire = rnd_date(dt.date(2018, 1, 1), dt.date(2026, 10, 20))
        leave, why = None, ""
        p_leave = 0.30
        if site == "測試社區C":
            p_leave = 0.55   # 埋：社區 C 離職集中
        if random.random() < p_leave and hire < dt.date(2026, 10, 1):
            lo = max(hire + dt.timedelta(days=20), dt.date(2025, 6, 1))
            if lo < dt.date(2026, 10, 31):
                leave = rnd_date(lo, dt.date(2026, 10, 31))
                if site == "測試社區C" and random.random() < 0.6:
                    leave = rnd_date(max(lo, dt.date(2026, 6, 1)), dt.date(2026, 10, 31))
                why = random.choice(reasons)
        if site == "測試社區E" and dt.date(2026, 3, 1) <= hire <= dt.date(2026, 8, 31) and random.random() < 0.7:
            leave = hire + dt.timedelta(days=random.randint(10, 80))   # 埋：社區 E 新人早離
            why = random.choice(["工時太長", "試用期未通過", "個人因素"])
        tags = []
        if comp == "甲保全":
            tags.append("保全")
        r = random.random()
        if r < 0.06:
            tags.append("急救人員")
        elif r < 0.10:
            tags.append("防火管理人")
        if title == "總幹事":
            tags.append("事務管理人員")
        if title == "機電技師":
            tags.append("設備安全管理人員")
        if n in (1, 151):
            tags.append("職安主管")
        eid = "E%04d" % n
        name = "範例員工%03d" % n
        people.append([eid, name, comp, site, title, hire, leave, why, "、".join(tags)])
        if leave and leave <= BASE:
            continue
        # 訓練紀錄
        if "保全" in tags and n % 37 != 0:      # 埋：少數保全沒有職前訓練紀錄
            trains.append([eid, name, "保全職前專業訓練", hire - dt.timedelta(days=7), 40, "", None])
        if n % 23 != 0:
            trains.append([eid, name, "一般安全衛生教育訓練（新僱）", hire, 3, "", None])
        if "保全" in tags:
            for month in (8, 9, 10):
                if month == 10 and n % 9 == 0:
                    continue                   # 埋：10 月在職訓練沒上
                if dt.date(2026, month, 28) < hire:
                    continue
                trains.append([eid, name, "保全在職訓練", dt.date(2026, month, random.randint(5, 25)), 4, "", None])
        if (BASE - hire).days > 365 * 2:
            last = rnd_date(BASE - dt.timedelta(days=365 * 3 + 120), BASE - dt.timedelta(days=30))
            trains.append([eid, name, "一般勞工安全衛生在職教育訓練", last, 3, "", None])
        if "急救人員" in tags:
            trains.append([eid, name, "急救人員在職訓練", rnd_date(dt.date(2023, 6, 1), dt.date(2025, 12, 1)), 3, "", None])
        if "防火管理人" in tags:
            trains.append([eid, name, "防火管理人初訓", rnd_date(dt.date(2021, 1, 1), dt.date(2023, 12, 31)), 12, "FM%05d" % n, None])
        if "職安主管" in tags:
            trains.append([eid, name, "職業安全衛生業務主管在職訓練", rnd_date(dt.date(2024, 6, 1), dt.date(2025, 3, 1)), 6, "", None])
        if "事務管理人員" in tags or "設備安全管理人員" in tags:
            got = rnd_date(dt.date(2021, 6, 1), dt.date(2024, 6, 1))
            trains.append([eid, name, "公寓大廈管理服務人員認可證", got, None, "AP%05d" % n, got.replace(year=got.year + 5)])
            for _ in range(random.randint(0, 2)):
                trains.append([eid, name, "公寓大廈管理服務人員回訓講習", rnd_date(got, BASE), 8, "", None])

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "員工名冊"
header(ws, ROSTER_COLS, {1: 14, 2: 16, 7: 14, 8: 22})
for p in people:
    ws.append(p)
for row in ws.iter_rows(min_row=2):
    for c in (row[5], row[6]):
        c.number_format = "yyyy/mm/dd"
ws2 = wb.create_sheet("訓練紀錄")
header(ws2, TRAIN_COLS, {2: 30, 3: 12, 6: 12})
for t in sorted(trains, key=lambda x: (x[0], x[3])):
    ws2.append(t)
for row in ws2.iter_rows(min_row=2):
    for c in (row[3], row[6]):
        c.number_format = "yyyy/mm/dd"
wb.save(os.path.join(HERE, "測試假資料_員工名冊與訓練紀錄.xlsx"))

# ── 員工名冊範本 ──
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "填寫說明"
for line in ["員工名冊範本：人才流動洞察、回訓人員管理都讀這個檔。",
             "「員工名冊」：一人一列。在職的人「離職日」留白。日期寫 2026/3/5 或 115/03/05 都可以。",
             "「職務標籤」：寫這個人需要哪些法定訓練，多個用「、」分開。可用：保全、急救人員、防火管理人、職安主管、"
             "事務管理人員、防火避難設施管理人員、設備安全管理人員。",
             "「訓練紀錄」：上完一次訓練就加一列。訓練項目請用下拉選單（名稱要一致才算得到）。",
             "證照類（公寓大廈管理服務人員認可證）請填「證照到期日」。",
             "⚠ 真實資料請存在 hr-tools 的 _local 資料夾，不會被上傳到 GitHub。"]:
    ws.append([line])
ws.column_dimensions["A"].width = 120
ws = wb.create_sheet("員工名冊")
header(ws, ROSTER_COLS, {1: 14, 2: 16, 7: 14, 8: 22})
ws.append(["E0001", "（範例）王小明", "某某保全", "某某社區", "保全員", dt.date(2024, 3, 1), None, "", "保全、急救人員"])
ws2 = wb.create_sheet("訓練紀錄")
header(ws2, TRAIN_COLS, {2: 30, 3: 12, 6: 12})
ws2.append(["E0001", "（範例）王小明", "保全在職訓練", dt.date(2026, 10, 15), 4, "", None])
dv = DataValidation(type="list", formula1='"' + ",".join(TRAININGS) + '"', allow_blank=True)
ws2.add_data_validation(dv)
dv.add("C2:C5000")
wb.save(os.path.join(HERE, "員工名冊_範本.xlsx"))

# ── 簽核紀錄（假資料＋範本）──
kinds = ["請假單", "加班申請單", "調班申請單", "補打卡單", "離職申請單", "制服領用單", "薪資異動單", "到職報到單"]
approvers = ["人資主管", "保全部經理", "物業部經理", "財務主管", "總經理"]
rows = []
for i in range(1, 46):
    comp, _, sites = random.choice(companies)
    applied = rnd_date(dt.date(2026, 10, 1), dt.date(2026, 10, 31))
    status = random.choice(["待簽核"] * 4 + ["已核准"] * 5 + ["已退回"])
    if i <= 6:
        applied = dt.date(2026, 10, random.randint(20, 26))   # 埋：超過 3 天未處理
        status = "待簽核"
    done = None if status == "待簽核" else applied + dt.timedelta(days=random.randint(0, 6))
    rows.append(["F%04d" % i, random.choice(kinds), "範例員工%03d" % random.randint(1, 300), comp, random.choice(sites),
                 applied, random.choice(approvers), status, done, ""])
rows.append(["F0046", "請假單", "範例員工010", "甲保全", "測試社區A", dt.date(2026, 10, 31), "保全部經理", "待簽核", None, "今天新進"])
for fname, data in [("測試假資料_簽核紀錄.xlsx", rows), ("簽核紀錄_範本.xlsx", [["F0001", "請假單", "（範例）王小明", "某某保全", "某某社區",
                                                                         dt.date(2026, 11, 2), "人資主管", "待簽核", None, ""]])]:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "簽核紀錄"
    header(ws, APPROVAL_COLS, {1: 14, 2: 16, 9: 24})
    for r in data:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in (row[5], row[8]):
            c.number_format = "yyyy/mm/dd"
    dv = DataValidation(type="list", formula1='"待簽核,已核准,已退回,已撤回"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("H2:H5000")
    wb.save(os.path.join(HERE, fname))
print("ok", len(people), "people", len(trains), "trainings", len(rows), "approvals")
