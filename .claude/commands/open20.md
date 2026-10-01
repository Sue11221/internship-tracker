---
description: Open today's 20 internship application pages in a new Chrome window, every tab connected to Claude
argument-hint: "[new]  (optional: build another batch of 20 first)"
---

Open today's list of 20 application pages in Chrome through the Claude in Chrome
extension, so that every tab sits in this session's tab group and can be driven
by Claude afterwards.

Arguments: $ARGUMENTS

## Steps

1. Get the links. In the project folder run:

       PYTHONIOENCODING=utf-8 .venv/Scripts/python.exe shortlist.py --urls

   - If the argument is `new`, first run `shortlist.py --new` (another batch of 20).
   - If it reports there is no list for today, run `shortlist.py` (no arguments) and
     then `--urls` again. That only re-picks from data already on disk. If the
     "data fetched" time it prints is more than a day old, say so and offer to run
     `update.bat quiet` (about 20 minutes) instead of doing it unasked.

2. Load the Chrome tools in ONE ToolSearch call:
   `select:mcp__claude-in-chrome__tabs_context_mcp,mcp__claude-in-chrome__tabs_create_mcp,mcp__claude-in-chrome__navigate,mcp__claude-in-chrome__browser_batch,mcp__claude-in-chrome__tabs_close_mcp`

3. Call `tabs_context_mcp` with `createIfEmpty: true`. With no existing group this
   opens a NEW Chrome window holding a new tab group with one empty tab. If the
   group already has tabs (an earlier list, maybe half-filled forms), leave them
   alone and just add the new ones.

4. One `browser_batch` of `tabs_create_mcp` calls so there is one empty tab per
   link (19 more when the group started with one empty tab). Collect the tab IDs.

5. One `browser_batch` of `navigate` calls, each with an explicit `tabId`, in list
   order (#1 in the first tab), ending with `tabs_context_mcp`. If the batch stops
   part-way (site permission, timeout), run the remaining links in a new batch;
   do not retry a failing link more than twice.

6. Verify with the final `tabs_context_mcp` output that every link has a tab in
   the group, then report a short table: #, company, page title, loaded or not.
   LEAVE ALL TABS OPEN. They are the deliverable.

## Rules

- Opening pages only. Do not fill, upload, or submit anything unless the user
  asks for that separately.
- When asked to help apply: fill fields and attach the matching resume PDF from
  the project folder (Quant, SWE, or AI_Consulting), then stop for the user to
  review. The user presses Submit. Never create accounts or type passwords.
- Work authorization answers: authorized to work in the US = Yes; needs
  sponsorship now or in the future = No (green card holder).

## The applied log

Every role on a daily list is already in the applied log (`shortlists/applied log.xlsx`,
kept in `data/applied_log.json`) and will never be listed again. Roles whose
posting closes are removed from the log automatically on each run.

If the user says they did NOT apply to some of today's roles, take those out of
the log so they come back in a later list (numbers are the # column of today's
list):

    PYTHONIOENCODING=utf-8 .venv/Scripts/python.exe shortlist.py --skip 3 7
