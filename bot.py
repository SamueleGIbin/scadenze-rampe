import os
import datetime as dt
import requests
import openpyxl

def t(v):
    if isinstance(v, dt.datetime):
        return v.time()
    if isinstance(v, dt.time):
        return v
    return None

def build():
    ws = openpyxl.load_workbook("data.xlsm", data_only=True, read_only=True)["Outbound"]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        tour, dest, dep, frt, ramp, closed = r[1], r[4], t(r[6]), t(r[8]), r[22], r[23]
        if tour and isinstance(ramp, int) and 300 <= ramp < 400 and dep and not closed:
            rows.append((dep, ramp, tour, dest, frt))
    rows.sort()
    out = ["*RAMPE 300*"]
    for dep, ramp, tour, dest, frt in rows:
        line = f"R{ramp} | {dest} | {tour} | CHIUDE {dep:%H:%M}"
        if frt:
            line += f" | FRT entro {frt:%H:%M}"
        out.append(line)
    return "\n".join(out)

def send(text):
    body = [{"type": "TextBlock", "text": l, "wrap": True} for l in text.split("\n")]
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

if __name__ == "__main__":
    send(build())
