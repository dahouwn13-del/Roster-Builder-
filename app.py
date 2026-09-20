from __future__ import annotations

import json
import os
import re
import threading
import webbrowser
from copy import copy
from datetime import datetime, timedelta
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
TEMPLATE = ROOT / "Roster_Template.xlsx"
EXPORTS = ROOT / "exports"
EXCLUDED_ROLES = {"DOR", "DGE", "AFOM", "GSM", "NM", "ANM", "DM"}
EXCLUDED_NAMES = {"harshit dhawan", "tatiana shvets", "hashan manatungha", "robinson thimothy paul", "ma. bianca timoner", "merry fomenko", "walid dahou"}


def clean(value):
    return " ".join(str(value or "").replace("\n", " ").split()).strip()


def load_team():
    ws = load_workbook(TEMPLATE, data_only=False).active
    team, department = [], "Operations"
    for row in range(7, ws.max_row + 1):
        role, name = clean(ws.cell(row, 1).value), clean(ws.cell(row, 2).value)
        if not name:
            continue
        if role:
            department = role
        if role in EXCLUDED_ROLES or name.lower().strip() in EXCLUDED_NAMES:
            continue
        existing = [clean(ws.cell(row, col).value) for col in range(5, 12)]
        preferred = next((x for x in existing if re.search(r"\d{1,2}:\d{2}", x)), "08:00 - 17:30")
        employee_id = str(ws.cell(row, 3).value or "")
        team.append({
            "id": f"{employee_id}|{name.rstrip('* ').lower()}", "employeeId": employee_id, "name": name.rstrip("* "),
            "department": department, "phone": str(ws.cell(row, 4).value or ""),
            "preferred": normalize_shift(preferred), "pendingOff": pending_count(ws.cell(row, 12).value),
            "pendingPH": 0, "active": True,
        })
    return team


def pending_count(value):
    text = clean(value).lower()
    match = re.search(r"(\d+)\s*(?:pending\s*)?off", text)
    return int(match.group(1)) if match else (1 if "pending off" in text else 0)


def normalize_shift(value):
    text = clean(value).upper()
    if not re.search(r"\d{1,2}:\d{2}", text):
        return text or "OFF"
    start = re.findall(r"\d{1,2}:\d{2}", text)[0]
    h, m = map(int, start.split(":"))
    end = datetime(2000, 1, 1, h, m) + timedelta(hours=9, minutes=30)
    return f"{h:02d}:{m:02d} - {end:%H:%M}"


def shift_start(code):
    match = re.search(r"(\d{1,2}):(\d{2})", code or "")
    return int(match.group(1)) * 60 + int(match.group(2)) if match else None


def classify(code):
    minute = shift_start(code)
    if minute is None: return "leave"
    if minute >= 19 * 60 or minute < 5 * 60: return "night"
    if minute < 11 * 60: return "morning"
    return "evening"


def department_shifts(dept):
    d = dept.lower()
    if "nest" in d: return ["07:00 - 16:30", "12:00 - 21:30", "14:00 - 23:30"]
    if "bell" in d or "door" in d or "concierge" in d: return ["07:00 - 16:30", "14:00 - 23:30", "23:00 - 08:30"]
    return ["08:00 - 17:30", "14:00 - 23:30", "23:00 - 08:30"]


def demand_score(day):
    return float(day.get("occupancy", 0)) * 100 + float(day.get("arrivals", 0)) * .45 + float(day.get("departures", 0)) * .35 + float(day.get("groups", 0)) * .7


def generate(payload):
    team = [x for x in payload["team"] if x.get("active", True)]
    days = payload["days"]
    schedule = {x["id"]: [] for x in team}
    week_no = datetime.fromisoformat(payload["weekStart"]).isocalendar().week
    by_dept = {}
    for person in team: by_dept.setdefault(person["department"], []).append(person)

    # Mixed 2/1 OFF pattern; alternates for each employee on the following week.
    targets = {p["id"]: 2 if (i + week_no) % 2 == 0 else 1 for i, p in enumerate(team)}
    off_days = {}
    daily_offs = [0] * 7
    for i, p in enumerate(team):
        count = targets[p["id"]]
        # Balance OFFs across the week first, then favor quieter operating days.
        choices = []
        for _ in range(count):
            available = [d for d in range(7) if d not in choices]
            chosen = min(available, key=lambda d: (daily_offs[d], demand_score(days[d]), (d - i) % 7))
            choices.append(chosen); daily_offs[chosen] += 1
        off_days[p["id"]] = set(choices)

    vivek = next((p for p in team if p["name"].lower().startswith("vivek ray")), None)
    shubham = next((p for p in team if p["name"].lower().startswith("shubham gupta")), None)
    if vivek:
        v_off = min(range(7), key=lambda d: demand_score(days[d]))
        off_days[vivek["id"]] = {v_off}
        if shubham:
            off_days[shubham["id"]].discard(v_off)
            needed = targets[shubham["id"]] - len(off_days[shubham["id"]])
            replacements = [d for d in sorted(range(7), key=lambda x: demand_score(days[x])) if d != v_off and d not in off_days[shubham["id"]]]
            off_days[shubham["id"]].update(replacements[:needed])

    for day_idx, day in enumerate(days):
        score = demand_score(day)
        for dept, people in by_dept.items():
            working = [p for p in people if day_idx not in off_days[p["id"]]]
            for p in people:
                if day_idx in off_days[p["id"]]: schedule[p["id"]].append("OFF")
                else: schedule[p["id"]].append(None)
            shifts = department_shifts(dept)
            for j, p in enumerate(working):
                if vivek and p["id"] == vivek["id"]: code = "21:00 - 06:30"
                elif shubham and vivek and p["id"] == shubham["id"] and day_idx in off_days[vivek["id"]]: code = "21:00 - 06:30"
                else:
                    # More PM/night coverage on high arrivals, groups, or occupancy.
                    offset = 1 if score >= 90 else 0
                    code = shifts[(j + day_idx + offset) % len(shifts)]
                    if day_idx and classify(schedule[p["id"]][day_idx - 1]) == "night" and classify(code) == "morning":
                        code = shifts[1]
                schedule[p["id"]][day_idx] = code
    return {"schedule": schedule, "offTargets": targets}


def export_roster(payload):
    wb = load_workbook(TEMPLATE)
    ws = wb.active
    start = datetime.fromisoformat(payload["weekStart"])
    team_map = {x["name"].strip().lower(): x for x in payload["team"] if x.get("active", True)}
    schedule = payload["schedule"]
    # Remove merged ranges before row deletion; the operational group labels are
    # restored on every retained row so filtering/sorting remains safe in Excel.
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    rows_to_delete = []
    for row in range(7, ws.max_row + 1):
        name = clean(ws.cell(row, 2).value).rstrip("* ").lower()
        role = clean(ws.cell(row, 1).value)
        if role in EXCLUDED_ROLES or name in EXCLUDED_NAMES:
            rows_to_delete.append(row)
    for row in reversed(rows_to_delete): ws.delete_rows(row)
    for i in range(7):
        date = start + timedelta(days=i)
        ws.cell(1, 5 + i).value = date.strftime("%A")
        ws.cell(2, 5 + i).value = date
        ws.cell(3, 5 + i).value = payload["days"][i].get("arrivals", 0)
        ws.cell(4, 5 + i).value = payload["days"][i].get("departures", 0)
        ws.cell(5, 5 + i).value = float(payload["days"][i].get("occupancy", 0)) / 100
        ws.cell(6, 5 + i).value = payload["days"][i].get("remarks", "")
    for row in range(7, ws.max_row + 1):
        key = clean(ws.cell(row, 2).value).rstrip("* ").lower()
        person = team_map.get(key)
        if not person: continue
        codes = schedule.get(person["id"], [])
        for i in range(7): ws.cell(row, 5 + i).value = codes[i] if i < len(codes) else ""
        ws.cell(row, 1).value = person["department"]
        ws.cell(row, 12).value = f"Pending OFF: {person.get('pendingOff', 0)} | Pending PH: {person.get('pendingPH', 0)}"
    ws.title = f"{start:%d.%m} to {(start + timedelta(days=6)):%d.%m}"
    ws.print_area = f"A1:L{ws.max_row}"
    EXPORTS.mkdir(exist_ok=True)
    path = EXPORTS / f"CIEL_Roster_{start:%Y-%m-%d}.xlsx"
    wb.save(path)
    return path


class Handler(SimpleHTTPRequestHandler):
    def translate_path(self, path):
        rel = urlparse(path).path.lstrip("/") or "index.html"
        return str(STATIC / rel)

    def do_GET(self):
        if self.path == "/api/team": return self.send_json(load_team())
        if self.path.startswith("/downloads/"):
            file = EXPORTS / Path(self.path).name
            if file.exists():
                self.send_response(200); self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{file.name}"'); self.end_headers(); self.wfile.write(file.read_bytes()); return
        return super().do_GET()

    def do_POST(self):
        size = int(self.headers.get("Content-Length", 0)); payload = json.loads(self.rfile.read(size) or b"{}")
        if self.path == "/api/generate": return self.send_json(generate(payload))
        if self.path == "/api/export":
            path = export_roster(payload); return self.send_json({"url": f"/downloads/{path.name}", "name": path.name})
        self.send_error(404)

    def send_json(self, obj):
        data = json.dumps(obj).encode(); self.send_response(200); self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data))); self.end_headers(); self.wfile.write(data)


if __name__ == "__main__":
    os.chdir(STATIC)
    url = "http://127.0.0.1:8765"
    threading.Timer(1, lambda: webbrowser.open(url)).start()
    print(f"CIEL Roster Generator running at {url}")
    ThreadingHTTPServer(("127.0.0.1", 8765), Handler).serve_forever()
