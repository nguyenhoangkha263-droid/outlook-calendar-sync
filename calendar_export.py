import win32com.client
from datetime import datetime, timedelta, UTC
from pathlib import Path
import uuid

OUTPUT_FILE = r"calendar.ics"

# number of days to export into the future
DAYS_AHEAD = 30

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# 9 = Calendar
calendar = namespace.GetDefaultFolder(9)

start = datetime.now()
end = start + timedelta(days=DAYS_AHEAD)

restriction = (
    "[Start] >= '{}'"
    " AND [Start] <= '{}'"
).format(
    start.strftime("%m/%d/%Y %H:%M %p"),
    end.strftime("%m/%d/%Y %H:%M %p")
)

items = calendar.Items
items.Sort("[Start]")
items.IncludeRecurrences = True

appointments = items.Restrict(restriction)

lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Outlook Sync//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
]

def ics_escape(text):
    if not text:
        return ""

    return (
        str(text)
        .replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\r", "")
        .replace("\n", "\\n")
    )

for appt in appointments:

    try:
        start_dt = appt.Start
        end_dt = appt.End

        start_str = start_dt.strftime("%Y%m%dT%H%M%S")
        end_str = end_dt.strftime("%Y%m%dT%H%M%S")

        subject = str(appt.Subject or "")
        location = str(appt.Location or "")
        organizer = str(appt.Organizer or "")
        body = ""

        try:
            body = str(appt.Body or "")
        except Exception:
            pass

        try:
            uid = str(appt.GlobalAppointmentID)
        except:
            uid = str(appt.EntryID)

        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{datetime.now(UTC).strftime('%Y%m%dT%H%M%SZ')}",
            f"DTSTART:{start_str}",
            f"DTEND:{end_str}",
            f"SUMMARY:{ics_escape(subject)}",
            f"LOCATION:{ics_escape(location)}",
            # f"ORGANIZER:{ics_escape(organizer)}",
            # f"DESCRIPTION:{ics_escape(body)}",
            "END:VEVENT"
        ])

    except Exception as ex:
        print("Skip:", ex)

lines.append("END:VCALENDAR")

Path(OUTPUT_FILE).write_text(
    "\r\n".join(lines),
    encoding="utf-8"
)

print(f"ICS exported: {OUTPUT_FILE}")