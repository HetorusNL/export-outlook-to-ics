"""Outlook ICS exporter, for sharing free/busy information with your other calendars

Optional environment variables to SCP the calendar file:
- SCP_HOST
- SCP_DST
"""

import os
from pathlib import Path
from win32com.client import Dispatch
from win32com.client import CDispatch

# OlDefaultFolders
# https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
olFolderCalendar = 9

# Calendar detail
olFullDetails = 2
olFreeBusyAndSubject = 1
olFreeBusyOnly = 0

# constants
CALENDER_FILE = "calendar.ics"


# stole the Outlook ICS exporter from Tom Smeets:
# https://github.com/TomSmeets/export-outlook-to-ics
# Copyright 2023 - Tom Smeets <tom@tsmeets.nl>
def export_outlook_calendar_to_ics(ics_file_path: str):
    outlook_app: CDispatch = Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get the default calendar folder
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getdefaultfolder
    folder = outlook_app.GetDefaultFolder(olFolderCalendar)

    # Export to ICS
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.calendarsharing.saveasical
    exporter = folder.GetCalendarExporter()
    exporter.CalendarDetail = olFreeBusyAndSubject
    exporter.IncludeAttachments = False
    exporter.IncludePrivateDetails = False
    exporter.IncludeWholeCalendar = True
    exporter.SaveAsICal(ics_file_path)


def scp_calendar_file():
    SCP_HOST = os.getenv("SCP_HOST")
    SCP_DST = os.getenv("SCP_DST")
    if not SCP_HOST or not SCP_DST:
        print(f"SCP_HOST and/or SCP_DST not set, skipping copying calendar file!")
        return

    command = f'wsl.exe -e bash -c "scp {CALENDER_FILE} {SCP_HOST}:{SCP_DST}"'
    print(f"running command: '{command}'")
    os.system(command)


if __name__ == "__main__":
    print("Start exporting ICS calendar")
    # Store the .ics file relative to this main file
    cwd = Path(__file__).parent.resolve(strict=True)
    ics_file_path = cwd / CALENDER_FILE

    export_outlook_calendar_to_ics(str(ics_file_path))
    print(f"Exported to:{ics_file_path}")

    scp_calendar_file()
