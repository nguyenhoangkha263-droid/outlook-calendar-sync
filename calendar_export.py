import win32com.client
from datetime import datetime, timedelta, UTC
from pathlib import Path
import uuid

OUTPUT_FILE = r"calendar.ics"

# number of days to export into the future and the past
DAYS_BACK = 7
DAYS_AHEAD = 30

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# 9 = Calendar
calendar = namespace.GetDefaultFolder(9)

# start = datetime.now() - timedelta(days=DAYS_BACK)
# end = start + timedelta(days=DAYS_BACK + DAYS_AHEAD)
start = (
        datetime.now() - timedelta(days=DAYS_BACK)
        ).replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )

end = (
        datetime.now() + timedelta(days=DAYS_AHEAD)
        ).replace(
            hour=23,
            minute=59,
            second=59,
            microsecond=999999
        )

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

        try:
            base_uid = str(appt.GlobalAppointmentID)
        except Exception:
            base_uid = str(appt.EntryID)

        occurrence_key = start_dt.strftime("%Y%m%dT%H%M%S")
        uid = f"{base_uid}-{occurrence_key}@outlook-calendar-sync"

        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTART:{start_str}",
            f"DTEND:{end_str}",
            f"SUMMARY:{ics_escape(subject)}",
            f"LOCATION:{ics_escape(location)}",
            "END:VEVENT",
        ])

    except Exception as ex:
        print("Skip:", ex)


lines.append("END:VCALENDAR")

old_content = ""

if Path(OUTPUT_FILE).exists():
    old_content = Path(OUTPUT_FILE).read_text(encoding="utf-8")

new_content = "\r\n".join(lines)

if new_content != old_content:
    Path(OUTPUT_FILE).write_text(new_content, encoding="utf-8")
    print("Calendar updated")
else:
    print("No changes")

print(f"ICS exported: {OUTPUT_FILE}")