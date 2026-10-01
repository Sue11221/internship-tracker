# Local setup notes

This folder is a clone of
https://github.com/zshah101/Automated-List-Of-Summer-2027-and-Fall-2026-Tech-Internships
with a Python virtualenv in `.venv`, plus a few personal scripts on top.

## The daily routine

1. Open `shortlists\<today's date>.xlsx`. It holds **20 roles at 20 different
   companies**, each with its application link in the **Apply link** column.
2. In Claude Code, opened in this folder, type `/open20` (or just say "open all
   20"). Claude opens a new Chrome window with all 20 application pages as tabs,
   connected to the Claude in Chrome extension so it can help fill them in.
3. Apply. You review every form and press Submit yourself.

Want more the same day? `.venv\Scripts\python.exe shortlist.py --new` builds
another 20 (`<date> batch 2.xlsx`), or say `/open20 new`.

## What is on the list

Roles come from every open posting in four areas (quant, data science / ML,
AI consulting, internet backend). One role per company per list. Order:

1. Manhattan / Jersey City in person, preferred dates
2. remote, preferred dates
3. Manhattan / Jersey City, other or unstated dates
4. other city in person, preferred dates
5. everything else (remote first, then other cities)

"Preferred dates" means the posting itself states dates from **June 1** on
(and inside the hard window below). "Other or unstated dates" means a stated
start between May 12 and May 31, or no dates stated at all (most postings).
Inside each step the four areas take turns, stated dates before unstated,
newest first. The **Priority** column of the daily Excel shows the step.

Dropped before the pick:

- roles that need US citizenship or a security clearance (a green card is not
  enough for these). Roles open to "citizens or permanent residents" are KEPT
  and carry the note "citizens or green card OK".
- PhD / Master's / MBA-only roles. Checked twice: by the title, and by the
  posting text ("currently pursuing a Master's or PhD", "Bachelor's degree,
  seeking a Master's"). A role open to bachelor's students but saying a PhD
  is preferred stays on the list with the note "PhD / Master's preferred,
  undergrads considered".
- roles whose posting states a graduation window that May 2028 misses
  (for example "graduating in 2027")
- postings the job board has already taken down. The tracker keeps a posting
  open until it misses it twice, so it can lag a day or two.
- terms that have already started (for example Fall 2026)
- anything not confirmed as a **summer** internship: the title, the tracker's
  season label, or the posting text has to say summer
- roles whose stated dates fall outside **May 12 - August 31**. This is a
  hard cut: a stated start of May 11 or earlier ("early May" included) is
  dropped, a stated end after August 31 is dropped, and a start given only as
  "May" with no day is dropped because it cannot be confirmed.

The **Summer dates** column shows the dates when the posting states them.
`term_dates.py` re-reads every candidate posting once a day (dates, summer
wording, degree level, graduation window, still up or not) and keeps the
result in `data/term_cache.json`. A posting it cannot read keeps yesterday's
facts; one never read at all gets the note "requirements not checked".
Your graduation month is `GRADUATION` at the top of `term_dates.py`. The window is set by `WINDOW_START`,
`WINDOW_END` and `PREFERRED_START` at the top of `shortlist.py`.

Postings of any age are used as long as they are still open.

## The applied log

`shortlists\applied log.xlsx` lists every role that has been on a daily list.

- A role in the log is never put on a list again.
- On each run, a logged role whose posting has closed (expired, or no longer
  an internship) is removed from the log.
- The log assumes you applied to everything on each list. If you skipped some,
  take them out so they come back later, using the # column of today's list:
  `.venv\Scripts\python.exe shortlist.py --skip 3 7` (or tell Claude
  "I skipped 3 and 7").

The log itself is `data\applied_log.json`; the Excel file is a readable copy
rebuilt on every run, so edits made inside the Excel file are not kept.

## What runs every day

A Windows scheduled task named **Internship Tracker Daily Update** runs
`update.bat quiet` at 10:00 AM (or as soon as the PC is next awake and on the
network). It needs you to be logged in; the window is hidden.

**The task is currently DISABLED.** To start the daily runs, open Task
Scheduler, find the task, and click Enable (or run in PowerShell:
`Enable-ScheduledTask -TaskName 'Internship Tracker Daily Update'`).

Steps of each run:

1. `run.py update`     fetch every company board (~20 min)
2. `term_dates.py`     read every candidate posting (dates, degree, still up)
   `shortlist.py`      prune the applied log, pick today's 20, write the Excel
3. `send_email.py`     email you today's 20 (needs the app password below)
4. `run.py notify`     Telegram/Discord alerts, no-op unless configured

Output of the last run: `logs\last_run.log`.
To run by hand, double-click `update.bat` (also opens the dashboard).
Running `shortlist.py` twice on the same day keeps the same list.
To change the time: Task Scheduler -> Task Scheduler Library -> the task -> Triggers.

## Email: one-time setup

The Gmail app password is never written to a file in this folder. It is kept
in Windows Credential Manager (Control Panel -> Credential Manager -> Windows
Credentials, entry `internship-tracker-gmail`), encrypted under your Windows
login. `email_config.json` holds only the addresses and server name.

1. Turn on 2-Step Verification for your Google account if it isn't already.
2. Go to https://myaccount.google.com/apppasswords, create one named
   "internship tracker", and copy the 16-letter code Google shows you.
3. Double-click `set_email_password.bat`. Paste the code (nothing is shown
   while you type), press Enter. It checks the code with Gmail, saves it in
   Credential Manager, then sends you a test email with the current list.

Until step 3 is done, the daily run simply skips the email step.
To change or remove the code later: run `set_email_password.bat` again, or
`.venv\Scripts\python.exe send_email.py --clear-password`. Deleting the app
password on Google's page also cuts the script off instantly.

After that, each daily run emails that day's 20. Set `"only_nyc": true` in
`email_config.json` to receive only New York area postings.

## Files

- `shortlists\2026-9-30.xlsx`     one workbook per day, named by the local date
  (no zero padding, dashes because Windows forbids slashes). Workbooks older
  than 5 days are deleted automatically.
- `shortlists\applied log.xlsx`   the applied log (never deleted)
- `data\today20.json`             today's 20, read by `/open20`
- `data\applied_log.json`         the applied log itself
- `.claude\commands\open20.md`    the `/open20` command
- `docs/index.html`               the tracker's own dashboard (every tech internship)
- `docs/internships.csv`, `docs/api/jobs.json`  raw tracker output

## Changing what it looks for

`shortlist.py` holds the settings near the top: `PER_DAY` (20), the keyword
regexes for quant / data science / backend, a list of consulting firms, a list
of internet companies, and the eligibility rules (`RE_GRAD_ONLY` and so on).

`data/config.json` controls the tracker itself:

- `regions`: `["US"]` (default), `["US", "Canada"]`, or `["Global"]`.
- `cycles`: which seasons to keep, e.g. `["Summer 2027", "Fall 2026"]`.
- `role_scope`: `"tech"` (software, data, ML, quant, hardware) or `"all"`.
- `max_per_company`: cap per employer (default 3).

The `supabase_*` keys in that file belong to the original author's public
email signup and are harmless locally.

## Local change to the tracker's own code

`src\intern_engine\sponsorship.py` was edited (classifier VERSION 5). Upstream
used one label, `citizens-only`, for two different things. Here it is split:
`citizens-only` = citizenship or a clearance is required; `us-persons` =
"citizens or permanent residents" / export-control "U.S. persons", which a
green card satisfies. `git pull upstream main` may conflict in that file.

## Adding companies

Add a slug to `data/candidates.json` and run
`.venv\Scripts\python.exe run.py harvest`. It probes Greenhouse, Lever,
Ashby, Workday, etc. to find which one hosts that company.
`run.py discover` grows the list automatically from public datasets.

## Your private copy on GitHub

https://github.com/Sue11221/internship-tracker (private) is the git remote
`origin`. To back up your changes:

    git add -A
    git commit -m "what changed"
    git push

GitHub Actions are turned off on it, so the original author's scheduled
workflows never run on your account. Local-only files stay off GitHub (see
`.gitignore`): the applied log, today's list, the daily Excel files, logs and
`email_config.json`. The email password is never in any file.

## Keeping up with the author's changes

The original public repo is the git remote `upstream`:

    git pull upstream main
    git push
