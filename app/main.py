import sys
import os
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse

from app.services.timesheet_comparator import process_timesheets

app = FastAPI()


@app.post("/compare-timesheets/")
async def compare_timesheets(
    sap_file: UploadFile = File(...),
    wand_file: UploadFile = File(...),
    mapping_file: UploadFile = File(...)
):
    output = await process_timesheets(sap_file, wand_file, mapping_file)
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={"Content-Disposition": "attachment; filename=comparison_report.xlsx"}
    )
