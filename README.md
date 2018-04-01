# leave-calendar

A leave calendar system that I implemented for my ex-team when i was learning how to use Google App Scripts. Also uses Google Forms, Sheets and Calendar

*Disclaimer: This is a fun project done when i was new to programming, please dont judge me on this!*

## the "architecture"

The leave calendar's "frontend" is a Google form that allows users to enter their names, leave's start date, end date, reason, and if they are creating a new entry or deleting a previous leave record.

The backend "Database" is on Google Sheets. It is linked to the form and records all responses there.

The backend "logic" is on Google App Script. It is linked to the sheet and is triggered on form submit.

## how to replicate

1. Create a new Google Calendar
    1. Go to *Settings and Sharing* and get its Calendar ID
    2. Should look something like this ocvqlkmnnadl5mfpauuk2v34mg@group.calendar.google.com
2. Create a form like [this](https://docs.google.com/forms/d/e/1FAIpQLSdx-mgberirAPbiL7_0Ef4O4AsHJcMGCL-lHAaVdqqM6wRtaA/viewform?usp=sf_link)
3. Create spreadsheet on the *RESPONSES* tab of the form to link it
    1. Note down the Sheet ID from the url of the Google Sheet e.g. https://docs.google.com/spreadsheets/d/1JiyR3mgaekPijSsW8UGk-HJayYLBPkesOqKVYv3QD2g/edit#gid=308592446
    2. The Sheet ID should then be: *1JiyR3mgaekPijSsW8UGk-HJayYLBPkesOqKVYv3QD2g*
4. Edit the spreadsheet so that it looks like [this](https://docs.google.com/spreadsheets/d/1JiyR3mgaekPijSsW8UGk-HJayYLBPkesOqKVYv3QD2g/edit?usp=sharing), with 2 sheets named records and events. The columns headers and order are extremely important
5. Go to *Tools* and click on *Script Editor*. This will open up the Google App Script editor.
6. Create a new script file and copy paste the entire Code.gs into it
    1. Update CALENDAR_ID with Calendar ID from step 1
    2. Update SPREADSHEET_ID with Sheet ID from step 3
7. On current project triggers, set
    1. Run - *onSubmit*
    2. Events - *From spreadsheet* and *On form submit*
8. Done!
