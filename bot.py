import os
import time
import datetime as dt
from zoneinfo import ZoneInfo
import requests
import openpyxl

TZ = ZoneInfo("Europe/Rome")
MAX = 345 * 60
COLORS = {1: "Good", 2: "Warning", 3: "Attention"}

def clock():
    return dt.datetime.now(TZ)

def secs(v):
    if isinstance(v, (dt.datetime, dt.time)):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, dt.timedelta):
        return int(v.total_seconds()) % 86400
    return None

def nxt(s, now):
    d = now.replace(hour=0, minute=0, second=0, microsecond=0) + dt.timedelta(seconds=s)
    return d if d >= now else d + dt.timedelta(days=1)

def minutes(dest, pre):
    tok = str(dest or "").upper().split()
    if "KG4" in tok or "DZ5" in tok:
        return 130
    if "MXP" in tok:
        return 50
    return 30 if pre == "ITD" else 55

def load(now):
    ws = openpyxl.load_workbook("data.xlsx", data_only=True).active
    ts = ws["H2"].value
    tasks = []
    for r in ws.iter_rows(min_row=2, max_col=6, values_only=True):
        ramp, dest, dep, frt, closed, pre = r[0], r[1], secs(r[2]), r[3], r[4], r[5]
        if not (isinstance(ramp, (int, float)) and dep is not None and not closed):
            continue
        p = nxt(dep, now)
        d = p - dt.timedelta(minutes=minutes(dest, pre))
        text = f"Baia {int(ramp)} - {dest} - da chiudere {d:%H:%M} ({p:%H:%M})"
        if frt:
            text += " - chiamare CR per girare freight"
        tasks.append({
            "deadline": d,
            "dep": p,
            "text": text,
            "zone": "BASSA" if ramp < 332 else "ALTA",
            "ramp": int(ramp),
            "dest": dest,
            "frt": bool(frt),
            "sent": 0,
        })
    tasks.sort(key=lambda t: t["deadline"])
    return ts if isinstance(ts, dt.datetime) else None, tasks

def step(tasks, now):
    out = []
    for t in tasks:
        m = (t["deadline"] - now).total_seconds() / 60
        lvl = 3 if m <= 10 else 2 if m <= 15 else 1 if m <= 30 else 0
        if m > 0 and lvl > t["sent"]:
            t["sent"] = lvl
            out.append((COLORS[lvl], t["text"], t["zone"]))
    return out

def send(items):
    body = []
    for c, x, z in items:
        if z:
            body.append({"type": "TextBlock", "text": z, "color": c, "weight": "Bolder", "size": "Small", "separator": True, "spacing": "Medium"})
        body.append({"type": "TextBlock", "text": x, "color": c, "weight": "Bolder", "wrap": True, "spacing": "None" if z else "Default"})
    requests.post(
        os.environ["TEAMS_URL"],
        json={
            "type": "message",
            "attachments": [{
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.4",
                    "body": body,
                },
            }],
        },
        timeout=30,
    ).raise_for_status()

def run():
    t0 = clock()
    ts, tasks = load(t0)
    tasks = [t for t in tasks if 0 < (t["deadline"] - t0).total_seconds() <= MAX]
    if not tasks:
        send([("Attention", "ATTENZIONE: nessuna scadenza nelle prossime 5h45. Controlla il file.", None)])
        return
    msg = f"BOT ATTIVO - {len(tasks)} scadenze - la prima alle {tasks[0]['deadline']:%H:%M}"
    if ts:
        msg += f" - file delle {ts:%H:%M}"
    items = [("Accent", msg, None)]
    for z in ("BASSA", "ALTA"):
        lst = ", ".join(f"{t['ramp']} {t['dest']} {t['deadline']:%H:%M} ({t['dep']:%H:%M})" for t in tasks if t["frt"] and t["zone"] == z)
        items.append(("Accent", f"DA GIRARE FREIGHT {z}: {lst or 'nessuna'}", None))
    send(items)
    zones = {t["zone"] for t in tasks}
    done = set()
    while tasks and (clock() - t0).total_seconds() < MAX:
        now = clock()
        out = step(tasks, now)
        tasks = [t for t in tasks if t["deadline"] > now]
        for z in ("BASSA", "ALTA"):
            if z in zones and z not in done and not any(t["zone"] == z for t in tasks):
                done.add(z)
                out.append(("Accent", f"FINE CHIUSURE {z}", None))
        if out:
            send(out)
        if tasks:
            time.sleep(30)
    if tasks:
        send([("Attention", "BOT FERMATO: limite di tempo, ricarica il file.", None)])

if __name__ == "__main__":
    run()
