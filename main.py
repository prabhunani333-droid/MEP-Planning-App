from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
from openpyxl import Workbook

app = FastAPI(title="MEP Planning & Monitoring App")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------- MODELS ----------------
class Activity(BaseModel):
    activity_id: str
    activity_name: str
    system: str
    duration: int
    start_day: int
    manpower: int
    planned_qty: float
    actual_qty: float

class ProjectData(BaseModel):
    activities: List[Activity]

# ---------------- EXPORT ----------------
@app.post("/export/excel")
def export_excel(data: ProjectData):
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Schedule"
    ws1.append([
        "Activity ID","Activity Name","System",
        "Start Day","Duration","Manpower"
    ])

    ws2 = wb.create_sheet("Progress")
    ws2.append([
        "Activity ID","Planned Qty","Actual Qty"
    ])

    for a in data.activities:
        ws1.append([
            a.activity_id, a.activity_name,
            a.system, a.start_day,
            a.duration, a.manpower
        ])
        ws2.append([
            a.activity_id,
            a.planned_qty,
            a.actual_qty
        ])

    wb.save("MEP_Planning_Output.xlsx")
    return {"status": "Excel exported successfully"}

@app.get("/")
def home():
    return {"message": "MEP Planning App Running"}
