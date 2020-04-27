#! python3
"""
Uses data from team schedule worksheet and creates calendar events on the google calendar,
which is then synced to the TOC website via Cloversites built-in sync function
"""

import toc
import kcal

tocSchedule = toc.schedule()
tocCalData = tocSchedule.schedule_data
tocCalendar = kcal.gcal(['TOC Test'])
tocCalendar.delete_duplicate_events(tocCalData)
tocCalendar.update_calendar(tocCalData)
tocSchedule.email_changes()