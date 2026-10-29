# Job Application Tracker

A small desktop app for logging job applications into an Excel tracker, built for job seekers in Switzerland who need to document their search effort for the RAV (regional employment office), but useful for anyone.

It asks you one question at a time, writes the row into your Excel file, creates a folder for the application, and celebrates your progress along the way.

## Features

* **Guided new entry**: one question per screen with sensible defaults, then a summary to check before saving. Missing details the RAV needs (address, contact person, email) are highlighted.
* **Automatic folders**: each application gets its own folder (name pattern configurable), optionally prefilled with copies of your CV and motivation letter templates, and the folder opens right away.
* **Edit and add entries**: searchable list of all applications with a full form to change any field.
* **Interview and job offer buttons**: pick the application, the tracker updates itself (status, last update, RAV result).
* **Monthly goal**: progress bar on the home screen that keeps growing past the goal, with 20 escalating bonus tiers.
* **Follow up reminders**: shows applications without news for a configurable number of days.
* **RAV export**: exports a month's applications in the column order of the official form *Nachweis der persönlichen Arbeitsbemühungen*.
* **Celebrations**: confetti, fireworks and sounds when you reach your goal, go beyond it, get an interview or receive an offer. Click anywhere to skip, or mute with one tap.
* Light, dark or system theme. All settings are stored in `settings.json`, no code editing needed.

## Requirements

* Windows (macOS and Linux work too, without the transparent celebration overlay)
* Python 3.7 or newer, for example via Anaconda
* Packages: `customtkinter`, `openpyxl`, `pillow`

```
python -m pip install customtkinter openpyxl pillow
```

## Setup

The app expects your tracker Excel file in the folder **above** the app folder:

```
Job_market/
    Application_Tracker.xlsx
    12_11_Company/                 (created per application)
    RAV_exports/                   (created by the export)
    Job_app_tracker/
        job_tracker.pyw
        assets/
```

1. Clone or download this repository into the folder that holds (or will hold) your tracker.
2. If you don't have a tracker yet, copy `template/Application_Tracker_template.xlsx` one level up and rename it as you like.
3. Start the app with `pythonw job_tracker.pyw` (no console window) or `python job_tracker.pyw` (shows errors).

If your tracker lives elsewhere, set the job folder and tracker file in Settings (gear icon).

### Desktop shortcut on Windows

Right click the desktop, choose New > Shortcut, and enter the full path to `pythonw.exe` followed by the full path to the script, both in quotes:

```
"C:\Path\To\Anaconda3\pythonw.exe" "C:\Path\To\Job_market\Job_app_tracker\job_tracker.pyw"
```

Run `where pythonw` in a terminal to find the first path.

## Tracker columns

Columns are matched by their header text in row 1, so the order doesn't matter and you can add your own columns. Recognised headers (matching the start of the header text):

| Header | Notes |
|---|---|
| ID | Filled automatically with the folder name |
| Date Applied | dd.mm.yyyy |
| Company | Required |
| Address | Needed for the RAV form |
| Contact Person | Needed for the RAV form |
| Contact email | Email or phone |
| Position | Required |
| Workload | e.g. 100% |
| Application Method | e.g. Online |
| Source | e.g. LinkedIn |
| Job ad link | Saved as a clickable link |
| Location | |
| Status | Applied, Interview, Offer, Rejected, ... |
| Last update | Set to today whenever you edit an entry |
| Result for RAV | Pending, Interview, Rejected, Hired |
| Folder | Clickable link to the application folder |
| Notes | |

## Good to know

* **Close the Excel file before saving** from the app, otherwise Windows locks it. The app tells you if it is open.
* Before each save, a backup `.bak` copy of the tracker is written next to it.
* The sounds in `assets/` can be replaced with any `.wav` file via Settings.
* For the RAV in Switzerland: applications usually need to be submitted by the 5th of the following month, and the expected number per month is agreed with your RAV advisor. Check your canton's rules.
