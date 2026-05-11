"""
High School Management System API

A super simple FastAPI application that allows students to view and sign up
for extracurricular activities at Mergington High School.
"""

from fastapi import FastAPI, HTTPException, File, UploadFile
from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse
import os
import re
from pathlib import Path
from typing import List
import openpyxl

app = FastAPI(title="Mergington High School API",
              description="API for viewing and signing up for extracurricular activities")

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# Mount the static files directory
current_dir = Path(__file__).parent


def is_valid_email(email: str) -> bool:
    return bool(EMAIL_REGEX.match(email.strip().lower()))


def parse_emails_from_workbook(upload_file: UploadFile) -> List[str]:
    try:
        workbook = openpyxl.load_workbook(upload_file.file, data_only=True)
    except Exception as exc:
        raise HTTPException(
            status_code=400,
            detail="Unable to read the Excel file. Please upload a valid .xlsx file.",
        ) from exc

    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))

    if not rows:
        raise HTTPException(status_code=400, detail="The Excel file is empty.")

    header = rows[0]
    email_col = None
    for idx, value in enumerate(header):
        if isinstance(value, str) and value.strip().lower() == "email":
            email_col = idx
            break

    if email_col is None:
        start_row = 0
        email_col = 0
    else:
        start_row = 1

    emails = []
    for row in rows[start_row:]:
        if row is None:
            continue
        cell = row[email_col] if len(row) > email_col else None
        if cell is None:
            continue
        email = str(cell).strip()
        if email:
            emails.append(email)

    if not emails:
        raise HTTPException(status_code=400, detail="No email addresses found in the Excel file.")

    return emails
app.mount("/static", StaticFiles(directory=os.path.join(Path(__file__).parent,
          "static")), name="static")

# In-memory activity database
activities = {
    "Chess Club": {
        "description": "Learn strategies and compete in chess tournaments",
        "schedule": "Fridays, 3:30 PM - 5:00 PM",
        "max_participants": 12,
        "participants": ["michael@mergington.edu", "daniel@mergington.edu"]
    },
    "Programming Class": {
        "description": "Learn programming fundamentals and build software projects",
        "schedule": "Tuesdays and Thursdays, 3:30 PM - 4:30 PM",
        "max_participants": 20,
        "participants": ["emma@mergington.edu", "sophia@mergington.edu"]
    },
    "Gym Class": {
        "description": "Physical education and sports activities",
        "schedule": "Mondays, Wednesdays, Fridays, 2:00 PM - 3:00 PM",
        "max_participants": 30,
        "participants": ["john@mergington.edu", "olivia@mergington.edu"]
    },
    "Soccer Team": {
        "description": "Join the school soccer team and compete in matches",
        "schedule": "Tuesdays and Thursdays, 4:00 PM - 5:30 PM",
        "max_participants": 22,
        "participants": ["liam@mergington.edu", "noah@mergington.edu"]
    },
    "Basketball Team": {
        "description": "Practice and play basketball with the school team",
        "schedule": "Wednesdays and Fridays, 3:30 PM - 5:00 PM",
        "max_participants": 15,
        "participants": ["ava@mergington.edu", "mia@mergington.edu"]
    },
    "Art Club": {
        "description": "Explore your creativity through painting and drawing",
        "schedule": "Thursdays, 3:30 PM - 5:00 PM",
        "max_participants": 15,
        "participants": ["amelia@mergington.edu", "harper@mergington.edu"]
    },
    "Drama Club": {
        "description": "Act, direct, and produce plays and performances",
        "schedule": "Mondays and Wednesdays, 4:00 PM - 5:30 PM",
        "max_participants": 20,
        "participants": ["ella@mergington.edu", "scarlett@mergington.edu"]
    },
    "Math Club": {
        "description": "Solve challenging problems and participate in math competitions",
        "schedule": "Tuesdays, 3:30 PM - 4:30 PM",
        "max_participants": 10,
        "participants": ["james@mergington.edu", "benjamin@mergington.edu"]
    },
    "Debate Team": {
        "description": "Develop public speaking and argumentation skills",
        "schedule": "Fridays, 4:00 PM - 5:30 PM",
        "max_participants": 12,
        "participants": ["charlotte@mergington.edu", "henry@mergington.edu"]
    }
}


@app.get("/")
def root():
    return RedirectResponse(url="/static/index.html")


@app.get("/activities")
def get_activities():
    return activities


@app.post("/activities/{activity_name}/signup")
def signup_for_activity(activity_name: str, email: str):
    """Sign up a student for an activity"""
    # Validate activity exists
    if activity_name not in activities:
        raise HTTPException(status_code=404, detail="Activity not found")

    # Get the specific activity
    activity = activities[activity_name]

    # Validate student is not already signed up
    if email in activity["participants"]:
        raise HTTPException(
            status_code=400,
            detail="Student is already signed up"
        )

    # Add student
    activity["participants"].append(email)
    return {"message": f"Signed up {email} for {activity_name}"}


@app.post("/activities/{activity_name}/import")
async def import_activity_attendance(
    activity_name: str, file: UploadFile = File(...)
):
    if activity_name not in activities:
        raise HTTPException(status_code=404, detail="Activity not found")

    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Please upload a valid .xlsx file.")

    activity = activities[activity_name]
    emails = parse_emails_from_workbook(file)

    added = []
    skipped = []
    invalid = []

    for email in emails:
        normalized_email = email.lower()
        if not is_valid_email(normalized_email):
            invalid.append(normalized_email)
            continue

        if normalized_email in activity["participants"]:
            skipped.append({"email": normalized_email, "reason": "duplicate"})
            continue

        if len(activity["participants"]) >= activity["max_participants"]:
            skipped.append({"email": normalized_email, "reason": "activity full"})
            continue

        activity["participants"].append(normalized_email)
        added.append(normalized_email)

    return {
        "message": f"Imported {len(added)} student(s).",
        "added": added,
        "skipped": skipped,
        "invalid": invalid,
    }


@app.delete("/activities/{activity_name}/unregister")
def unregister_from_activity(activity_name: str, email: str):
    """Unregister a student from an activity"""
    # Validate activity exists
    if activity_name not in activities:
        raise HTTPException(status_code=404, detail="Activity not found")

    # Get the specific activity
    activity = activities[activity_name]

    # Validate student is signed up
    if email not in activity["participants"]:
        raise HTTPException(
            status_code=400,
            detail="Student is not signed up for this activity"
        )

    # Remove student
    activity["participants"].remove(email)
    return {"message": f"Unregistered {email} from {activity_name}"}
