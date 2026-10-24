"""HTTP API for appending daily work activities to the master workbook."""

from __future__ import annotations

import datetime as dt
import logging
import uuid
from contextlib import asynccontextmanager
from typing import Literal
from urllib.parse import quote

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel, Field

from .config import REMARKS, WORK_STATUS, Settings
from .storage import (
    LockTimeout,
    MasterStore,
    SaveFailed,
    SheetMissing,
    WorkbookUnavailable,
    find_submission,
    payload_hash,
    record_submission,
)
from .workbook import (
    COL_ACTIVITY,
    SHEET_NAME,
    DateAlreadyHasActivities,
    FixedValues,
    append_activities,
    clean_activity,
    entries_for_date,
    parse_blocks,
)

log = logging.getLogger("dwa")

MSG_UNAVAILABLE = (
    "Unable to access the master Excel workbook. No activities have been submitted. Please try again later."
)
MSG_SAVE_FAILED = "The activities could not be saved. Please retry. Your submission has not been confirmed."
MSG_EMPTY = "Please enter at least one work activity."
MSG_SUCCESS = "Activities submitted successfully! The master Excel workbook has been updated."


def error(status: int, code: str, message: str, **extra) -> JSONResponse:
    return JSONResponse(status_code=status, content={"code": code, "message": message, **extra})


class SubmissionIn(BaseModel):
    submission_id: uuid.UUID
    date: dt.date
    shift: Literal["A", "B"]
    activities: list[str] = Field(max_length=1000)
    confirm_existing: bool = False


def create_app(settings: Settings | None = None) -> FastAPI:
    settings = settings or Settings()
    store = MasterStore(settings)

    @asynccontextmanager
    async def lifespan(_app: FastAPI):
        store.initialize()
        yield

    app = FastAPI(title="Daily Work Activity", lifespan=lifespan, docs_url=None, redoc_url=None)
    app.state.store = store

    @app.exception_handler(RequestValidationError)
    async def _validation(_req: Request, exc: RequestValidationError):
        return error(422, "invalid_input", "Some of the submitted details are invalid.", details=str(exc.errors()))

    @app.get("/api/health")
    def health():
        return {"ok": True}

    @app.get("/api/workbook/status")
    def workbook_status():
        try:
            wb = store.read()
        except WorkbookUnavailable:
            return error(503, "workbook_unavailable", MSG_UNAVAILABLE)
        except SheetMissing as e:
            return error(500, "sheet_missing", str(e))
        blocks = parse_blocks(wb[SHEET_NAME])
        dated = [b for b in blocks if b.date]
        latest = max(dated, key=lambda b: b.header_row) if dated else None
        return {
            "ok": True,
            "file_name": settings.master_name,
            "daily_blocks": len(blocks),
            "latest_entry": {"date": latest.date.isoformat(), "shift": latest.shift} if latest else None,
        }

    @app.get("/api/dates/{date}")
    def date_entries(date: dt.date):
        try:
            wb = store.read()
        except WorkbookUnavailable:
            return error(503, "workbook_unavailable", MSG_UNAVAILABLE)
        except SheetMissing as e:
            return error(500, "sheet_missing", str(e))
        return {"date": date.isoformat(), "entries": entries_for_date(wb[SHEET_NAME], date)}

    @app.post("/api/submissions")
    def submit(body: SubmissionIn):
        activities = [a for a in (clean_activity(x) for x in body.activities) if a]
        if not activities:
            return error(422, "empty", MSG_EMPTY)
        if len(activities) > settings.max_activities:
            return error(422, "too_many", f"Please submit at most {settings.max_activities} activities at a time.")
        too_long = [i + 1 for i, a in enumerate(activities) if len(a) > settings.max_activity_length]
        if too_long:
            return error(
                422, "too_long",
                f"Activity {too_long[0]} is longer than {settings.max_activity_length} characters.",
            )
        if not dt.date(2000, 1, 1) <= body.date <= dt.date(2100, 12, 31):
            return error(422, "invalid_date", "Please select a valid work date.")

        sid = str(body.submission_id)
        digest = payload_hash(body.date, body.shift, activities)
        fixed = FixedValues(
            cdl=settings.cdl_text, status=WORK_STATUS, dri=settings.shift_dri, remarks=REMARKS
        )
        try:
            with store.transaction() as tx:
                prior = find_submission(tx.wb, sid)
                if prior is not None:
                    # Retry of a request that was already saved (double click, network retry, refresh).
                    tx.save = False
                    if prior.get("h") != digest:
                        return error(409, "submission_conflict",
                                     "This submission id was already used for different activities. Please submit again.")
                    return {
                        "status": "duplicate",
                        "message": MSG_SUCCESS,
                        "added": prior.get("n", len(activities)),
                        "date": body.date.isoformat(),
                        "shift": body.shift,
                    }

                ws = tx.wb[SHEET_NAME]
                result = append_activities(
                    ws, body.date, body.shift, activities, fixed, body.confirm_existing
                )
                record_submission(tx.wb, sid, {
                    "d": body.date.isoformat(), "s": body.shift, "n": len(activities), "h": digest,
                    "r": [result.rows[0], result.rows[-1]],
                    "t": dt.datetime.now().isoformat(timespec="seconds"),
                })

                def verify(saved):
                    sheet = saved[SHEET_NAME]
                    got = [sheet.cell(r, COL_ACTIVITY).value for r in result.rows]
                    if got != activities:
                        raise SaveFailed("Saved activities do not match the submission")
                    if find_submission(saved, sid) is None:
                        raise SaveFailed("Submission record missing from saved workbook")

                tx.verify = verify
        except DateAlreadyHasActivities as e:
            return error(
                409, "date_exists",
                "Activities already exist for this date. Do you want to add more activities to this date?",
                entries=e.entries,
            )
        except (WorkbookUnavailable, LockTimeout) as e:
            log.error("workbook unavailable: %s", e)
            return error(503, "workbook_unavailable", MSG_UNAVAILABLE)
        except SheetMissing as e:
            log.error("sheet missing: %s", e)
            return error(500, "sheet_missing", f"{e}. No activities have been submitted. "
                                               "Please contact the administrator.")
        except Exception as e:  # SaveFailed or anything unexpected: never claim success
            log.exception("save failed: %s", e)
            return error(500, "save_failed", MSG_SAVE_FAILED)

        store.log({"id": sid, "date": body.date.isoformat(), "shift": body.shift,
                   "count": len(activities), "rows": result.rows,
                   "at": dt.datetime.now().isoformat(timespec="seconds")})
        return {
            "status": "saved",
            "message": MSG_SUCCESS,
            "added": len(activities),
            "date": body.date.isoformat(),
            "shift": body.shift,
            "rows": result.rows,
        }

    @app.get("/api/workbook/download")
    def download():
        try:
            data = store.read_bytes()
        except WorkbookUnavailable:
            return error(503, "workbook_unavailable", MSG_UNAVAILABLE)
        name = settings.master_name
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=\"master.xlsx\"; filename*=UTF-8''{quote(name)}",
                "Cache-Control": "no-store",
            },
        )

    return app


app = create_app()
