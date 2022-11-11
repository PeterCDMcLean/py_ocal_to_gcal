# Py Ocal to Gcal

# Python script to sync Outlook calendar Items to Google Calender

# Overview

This is a python script to synchronize an Outlook calendar to a separate person Google Calendar.

It is intended to be used for companies with restrictive security policies that prevent normal sharing of Outlook calendars with URLs.

If you prefer a GUI based solution, then there is [OWGC](https://github.com/phw198/OutlookGoogleCalendarSync) which is a .NET based open source solution

This script is intended for users who want a minimal solution that does not use .NET

There is an initial setup you must perform to enable your own 'Google Cloud' project which will allow you to authenticate the script with Google solely for
actions related to your calendar.

# Script Installation

Regrettably, because Outlook is only on Windows, this script only functions on Windows.

Note: I tested with 3.10.8. Other versions of Python may work for you.

Download Python 3.10.8 for Windows. Choose the installer.

[32 bit](https://www.python.org/ftp/python/3.10.8/python-3.10.8.exe)

[64 bit](https://www.python.org/ftp/python/3.10.8/python-3.10.8-amd64.exe)

We will create a virtual environment to run the script in
```

%LocalAppData%\Programs\Python\Python310\python -m venv venv

venv\Scripts\activate

pip install --upgrade pip

python -m pip install -r requirements.txt

```

Alternatively, you can use the install.bat script (This requires Python 3.10 in the PATH environment variable)

The first argument is a path to the venv:

```
install.bat venv
```


# Google Cloud Project Setup

1. Go to [Google Cloud Create Project](https://console.cloud.google.com/projectcreate)
  - Name your project what your want
  - Location can remain 'No organization'
  - Click Next
2. Click API Library (On the Left)
  - Search for Google Calendar
  - Click on Google Calendar API
  - Click Enable
3. Click on OAuth Consent Screen (On the Left)
  - User Type: 'External' 
  - We will only add your account as a tester, no other user will have access to this project if you do not add them or deploy
  - Click Create
4. App Registration
  - Type in an App Name
  - Add your email as the user support email
  - Add your email as the developer contact email (at the bottom)
  - Click Save and Continue
5. Scopes
  - Click Add or Remove Scopes
  - Filter for Google Calendar
    * If you do not see the Calendar API, you missed step 2. Please follow step 2. Refresh the scopes page
  - Enable .../auth/calendar.app.created 'Make secondary Google calendars, and see, create, change, and delete events on them '
  - Click Save and Continue
6. Test Users
  - Add your Google account as a test user.
  - Click Save and Continue
  - Click Back to Dashboard
7. Click Credentials (on the left)
  - Click Create Credentials
  - Select OAuth Client ID
  - Application Type 'Desktop App'
  - Click Download JSON
    * Save to C:\Users\<Username>\.credentials\credentials.json

# Running the program

From within the Python Virtual Environment

```
py_ocal_to_gcal.py --work [WorkName]
```

A calendar called [WorkName] will created if it does not exist. 

By default, this command will sync from 7 days ago to 31 days in the future to your soon to be create Calendar.

[WorkName] should be a unique calendar. Do not reuse a name of a Calendar you already have
If there is an existing Google Calendar matching [WorkName], **many events will be deleted!**

A web browser will open to Google to start the OAuth authentication process

NOTE: For some reason, the text Google shows does not match Step #5, even though we chose a very restricted Scope.

This will authenticate the script and allow it to access your Google Calendar

# Parting Words

You must guard this synchronized calendar with the same care as your other sensitive work information.

If you share your newly synchronized google calendar with other people, you may violate NDA by exposing meeting passcodes
or other critical important information.

Google does allow you to share only 'Free/Busy' calendar information with other Google users.
That may be sufficient to allow people to know your work hours without exposing company secrets. IANAL.
