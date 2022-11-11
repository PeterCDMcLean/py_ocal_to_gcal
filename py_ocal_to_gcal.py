'''
MIT License

Copyright (c) 2022 PeterCDMcLean

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''
import argparse
from datetime import date, datetime, timedelta
from gcsa.event import Event, Transparency
from gcsa.google_calendar import GoogleCalendar
from gcsa.calendar import AccessRoles, Calendar
import icalendar
from pathlib import Path
import pytz
import sys
import win32com.client

class CalendarHash:
  fromIsoDate = ""
  toIsoDate   = ""
  subject     = ""
  body        = ""
  location    = ""
  busy        = False # GCal: transparency : 'opaque' == Busy, 'transparent' = Free
  def __init__(self,
               fromIsoDate:str = "",
               toIsoDate:str = "",
               subject:str = "",
               body:str = "",
               location:str = "",
               busy:bool = False):
    self.fromIsoDate = fromIsoDate
    self.toIsoDate   = toIsoDate
    self.subject = subject
    self.body   = body
    self.location = location
    self.busy   = busy

  def __hash__ (self):
    return hash((
    self.fromIsoDate,
    self.toIsoDate  ,
    self.subject    ,
    self.body       ,
    self.location   ,
    self.busy       ))

  def __eq__ (self, b):
    if (( self.fromIsoDate == b.fromIsoDate ) and
        ( self.toIsoDate   == b.toIsoDate   ) and
        ( self.subject     == b.subject     ) and
        ( self.body        == b.body        ) and
        ( self.location    == b.location    ) and
        ( self.busy        == b.busy        )) :
       return True
    return False

  def print(self):
    print(f"{self.fromIsoDate} -> {self.toIsoDate}")
    print(f"{self.subject} @ {self.location}")
    print(f"{self.body}")
    print(f"busy {self.busy}")


def GetOutlookCalendar(fromUtc, toUtc):
  outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
  calendar = outlook.getDefaultFolder(9).Items
  calendar.IncludeRecurrences = True
  calendar.Sort('[Start]')

  restriction = "[Start] >= '" + fromUtc.strftime('%m/%d/%Y') + "' AND [End] <= '" + toUtc.strftime('%m/%d/%Y') + "'"
  calendar = calendar.Restrict(restriction)
  return calendar

def GetOutlookSet(fromUtc, toUtc):
  raw_cal = GetOutlookCalendar(fromUtc=fromUtc, toUtc=toUtc)
  outlook_cal_set = set()
  for e in raw_cal:
    outlook_item = CalendarHash(
      fromIsoDate = e.StartUTC.isoformat(),
      toIsoDate   = e.EndUTC.isoformat(),
      subject     = "" if e.Subject  is None else e.Subject.encode('ascii','ignore').decode('ascii'),
      body        = "" if e.Body     is None else e.Body.encode('ascii','ignore').decode('ascii'),
      location    = "" if e.Location is None else e.Location.encode('ascii','ignore').decode('ascii'),
      busy        = True if e.BusyStatus is None else e.BusyStatus == 2
    )

    outlook_cal_set.add(outlook_item)
  return outlook_cal_set

def ReadFreeBusy(filePath:Path) -> set:
  e = open(filePath, 'rb')
  ecal = icalendar.Calendar.from_ical(e.read())
  e.close()

  free_busy_set = set()

  for component in ecal.walk():
    if (component.name == "VFREEBUSY"):
      items = dict(component.items())
      for k in items['FREEBUSY']:
        kstart = k.start.astimezone(pytz.timezone('UTC'))
        kend = k.end.astimezone(pytz.timezone('UTC'))

        free_busy_set.add(
          CalendarHash(
            fromIsoDate = kstart.isoformat(),
            toIsoDate   = kend.isoformat(),
            subject     = "",
            body        = "",
            location    = "",
            busy        = True
          ))
  return free_busy_set


def FindWorkCalendar(work: str, gc:GoogleCalendar) -> Calendar:
  for cal in gc.get_calendar_list(
      min_access_role=AccessRoles.READER,
      show_deleted=True,
      show_hidden=True):
    if (cal.summary == work):
      return cal
  return None

def sync(before:int, after:int, work:str, email:str) -> int:

  today   = datetime.now()
  fromUtc = today - timedelta(days=before)
  toUtc   = today + timedelta(days=after)
  fromUtc = fromUtc.astimezone(pytz.timezone('UTC'))
  toUtc   = toUtc.astimezone(pytz.timezone('UTC'))

  #outlook_cal_set = ReadFreeBusy('freebusy.vfb')

  outlook_cal_set = GetOutlookSet(fromUtc, toUtc)

  gc = GoogleCalendar(email)

  workcal = FindWorkCalendar(work, gc)

  if (workcal is None):
    workcal = Calendar(
      work,
      description='Calendar Synchronized from ' + work)
    gc.add_calendar(workcal)
    workcal = FindWorkCalendar(work, gc)

  google_cal_set = set()
  google_cal_event_id_dict = dict()

  for event in gc.get_events(time_min=datetime(2000, 1, 1), time_max=datetime(3000,1,1), calendar_id=workcal.calendar_id):
    kstart = event.start.astimezone(pytz.timezone('UTC'))
    kend = event.end.astimezone(pytz.timezone('UTC'))
    #gc.delete_event(event.event_id, calendar_id=workcal.calendar_id)
    gcal_hash = CalendarHash(
                  fromIsoDate = kstart.isoformat(),
                  toIsoDate   = kend.isoformat(),
                  subject     = "" if event.summary     is None else event.summary,
                  body        = "" if event.description is None else event.description,
                  location    = "" if event.location    is None else event.location,
                  busy        = True if event.transparency is None else event.transparency == Transparency.OPAQUE
                )
    google_cal_set.add(gcal_hash)
    google_cal_event_id_dict[gcal_hash] = event.event_id


  missing_events = (outlook_cal_set-google_cal_set)
  if (len(missing_events) > 0):
    # Add missing events
    print('Add missing events')

    for missing in missing_events:
      kstart = datetime.fromisoformat(missing.fromIsoDate)
      kend = datetime.fromisoformat(missing.toIsoDate)
      print("Add")
      missing.print()
      trans = Transparency.OPAQUE
      if (not missing.busy):
        trans = Transparency.TRANSPARENT
      e = Event(start=kstart, end=kend, summary=missing.subject, description=missing.body, location=missing.location, transparency=trans)
      gc.add_event(e, calendar_id=workcal.calendar_id)

  extra_events = (google_cal_set-outlook_cal_set)
  if (len(extra_events) > 0):
    # Delete extra events
    print ('Delete extra events')
    for delthis in extra_events:
      print ("Delete")
      delthis.print()
      print(f"event id {google_cal_event_id_dict[delthis]}")
      gc.delete_event(google_cal_event_id_dict[delthis], calendar_id=workcal.calendar_id)
  return 0

if __name__ == '__main__':
  parser = argparse.ArgumentParser(
                      prog = 'py_cal_to_gcal',
                      description = 'Synchronize your Outlook Calendar to a Google Calendar.',
                      epilog = 'Please see the README.md file for detailed information.')

  parser.add_argument('-e', '--email', default=None, help='Google Account Email to Sync to.', required=True)
  parser.add_argument('-w', '--work', default='Work', help='Calendar in Google to Sync to.', required=True)
  parser.add_argument('-b', '--before', default=7, type=int, help='Days before current time to sync')
  parser.add_argument('-a', '--after' , default=31, type=int, help='Days after current time to sync')

  args = parser.parse_args()

  sys.exit(sync(before=args.before, after=args.after, work=args.work, email=args.email))


'''
OlBusyStatus enumeration (Outlook)
Name	Value	Description
olBusy	2	The user is busy.
olFree	0	The user is available.
olOutOfOffice	3	The user is out of office.
olTentative	1	The user has a tentative appointment scheduled.
olWorkingElsewhere	4	The user is working in a location away from the office.
'''

'''
Name 	Value 	Description
olFormatHTML 	2 	HTML format
olFormatPlain 	1 	Plain format
olFormatRichText 	3 	Rich text format
olFormatUnspecified 	0 	Unspecified format
'''