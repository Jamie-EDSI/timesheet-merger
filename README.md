# Contractor Timesheet Merger

Drop in a check report (PDF or Excel) and your ongoing Contractor Timesheet
workbook, and this tool merges the report's hours straight into the right
tabs, days, and fund columns - matched by Employee ID, one tab per person,
with every pay period stacked onto that same tab.

Everything runs **in your browser** - no backend, no login, nothing ever
uploaded anywhere. That's on purpose, and it's the same approach the
OSY/ISY Case Router this project is based on uses.

---

## What it does, step by step

1. You drop in the check report (the one broken out by Project Code, like
   `Timesheets_3_30__4_26_26_Berks.pdf`) and, if you have one, your existing
   ongoing timesheet workbook.
2. The app reads every daily hours row out of the report, figures out which
   2-week pay period each date falls into, and works out which fund each
   Project Code maps to (see below).
3. You get a review list of every employee found (name, ID, total hours) so
   you can uncheck anyone who shouldn't be included this run, then confirm.
4. For each included employee, it finds (or creates) their one tab, and
   writes each of their pay periods as its own block on it - reusing the
   *exact* template formatting, colors, and dropdown lists, not a re-built
   approximation - and fills in the manager field with their most frequent
   approver for that first period.
5. You get a single workbook back, with one tab per employee, ready to save
   as your new "ongoing" file for next time.

Run it again next pay period with the workbook you just downloaded as the
"ongoing" file, and it picks up right where it left off - it won't
duplicate a pay period it's already written.

---

## The Project Code → Fund mapping

This was the main unknown going in: the check report identifies hours by a
numeric **Project Code**, but the timesheet identifies them by **Fund
name** (a dropdown list). There's no field anywhere that connects the two,
so this mapping had to come from Morgan/EDSI directly. It lives in one
place: **`lib/fundMap.js`** - that's the file to edit if a new Project Code
shows up, or a fund name needs to change.

| Project Code | Fund written to the timesheet | Note |
|---|---|---|
| 525001 | Business Services | Added to the template's own BCWDB dropdown list (Q24 on Sheet1) so it's a normal, selectable option - not flagged. |
| 525002 | EARN | |
| 525006 | WIOA A | Normalized from "WIOA Adult" |
| 525007 | WIOA DW | |
| 525008 | OSO | Normalized from "One Stop Operator" |
| 525010 | *(PTO - Leave Time column, not a fund column)* | |
| 525011 | SNAP | |
| 525014 | ComOut | Normalized from "Community Outreach" |
| 525015 | VEAP | Normalized from "VEP Grant" (closest existing dropdown option) |

Any Project Code that isn't in this table at all gets written in as
`Unknown Code <code>` and flagged, rather than silently dropped.

---

## Known simplifications (read this before relying on it for real payroll)

This was built fast, as an MVP, and tested thoroughly - but a few
deliberate shortcuts are worth knowing about:

- **Only the first pay period on each employee's tab uses the template's
  original design fully** (rows 10-23, including its live formulas). Every
  pay period after that is a freshly-built block appended below, styled to
  match closely, but it's a close visual match rather than a pixel-perfect
  clone of every original border/shading detail. Where that next block
  starts is always computed dynamically - one row past whatever the highest
  row number already on the tab is - specifically so it can never land on
  top of the template's own EMPLOYEE SIGNATURE / MANAGER SIGNATURE lines
  (rows 28-29 and 34-35) or an earlier appended block. (An earlier version
  of this app used a fixed offset there, which is what caused the
  "cutting off hours" bug on multi-period tabs - two rows sharing the same
  row number, with one silently winning.)
- **Only 4 BCWDB fund columns and 3 Non-BCWDB fund columns exist per pay
  period** (that's what the template has room for). If an employee
  genuinely worked more than 4 distinct BCWDB funds in one 2-week period,
  the extra fund's hours get folded into the last column and flagged as a
  warning - it needs a human to look at it, the template just doesn't have
  more room.
- **Pay periods are inferred, not looked up from a real payroll calendar.**
  The app anchors to the Monday of the earliest date in the report and
  buckets everything into 14-day windows from there. If your actual pay
  period calendar doesn't start on that Monday, this is the place to fix
  it: `lib/payPeriods.js`.
- **Matching a report row to an employee's tab is by Employee ID**,
  normalized by stripping leading zeros (so `0000971` and `971` are treated
  as the same person - some report exports zero-pad IDs and some don't) and
  stored in a small marker cell (`A1`) on each tab that isn't part of the
  visible template. Don't delete or overwrite that cell, or the app won't
  recognize that tab on the next run. If a row's Name field has an Employee
  ID glued onto the front of it (e.g. `"1125662   Wilson, Amanda"` instead
  of separate ID/Name columns - seen once in a real report export), the
  parser strips it back down to a clean name and flags the row for review,
  rather than letting the garbled name create a phantom duplicate tab.
- **An "ongoing" workbook that still has tabs from before this A1 marker
  existed** (named just `"Lastname, First"`, no marker at all) is handled
  specially: if a fresh merge is about to write a tab under that same name,
  it replaces that legacy tab in place - clean template, fully written,
  properly marker-tagged - rather than leaving the old broken tab behind
  next to a newly-created, differently-named one (which is what used to
  happen, and looked like a duplicated/cut-off timesheet). A same-name
  collision between two *different*, already-tracked employees still falls
  back to an ID-suffixed tab name, same as always.
- **Holiday hours come from the report's Pay Code column** - a row with pay
  code `HOLIDAY` always lands in the Holiday column, regardless of what its
  Project Code would normally map to. PTO is still populated separately,
  since Project Code 525010 maps to it.
- **The MANAGER field is filled in with whichever approver shows up most
  often** for that employee in their first pay period (from the report's
  Approver column), with a trailing `(Terminated)` stripped off. Blank
  approvers and "Automatically Approved" don't count - if nothing else
  qualifies, the field is left blank for manual entry. It isn't re-derived
  for later pay periods appended to the same tab, since the template only
  has one MANAGER field per tab.
- **You get a review step before anything is written.** After the report is
  parsed, the app lists every employee it found (name, ID, total hours)
  with a checkbox so you can exclude anyone who shouldn't get a tab this
  run - selections aren't saved between runs.

None of these are hidden - the app surfaces a "Please review" list after
every run with anything that needs a second look.

---

## A note on `npm install` and security warnings

You may see `npm audit` flag a high-severity issue in the `xlsx` package
even after installing. This is expected and already handled deliberately:
SheetJS (the maintainers of `xlsx`) stopped publishing patched versions to
the public npm registry and now distribute fixes only from their own site,
so `package.json` points `xlsx` directly at their official patched
tarball (`cdn.sheetjs.com`) instead of the stale npm one. If `npm audit`
still shows something for `xlsx` after a fresh `npm install`, it means
`npm audit`'s advisory database doesn't recognize that install source (a
known quirk) rather than an unpatched install - you can confirm the actual
installed version with `npm ls xlsx`, which should show `0.20.3` or newer.

---

## Project structure

```
app/
  page.js              -- the UI (two drop zones, results table, warnings)
  layout.js / globals.css
lib/
  fundMap.js            -- Project Code -> Fund mapping (edit here first)
  reportParser.js        -- reads the PDF or Excel check report into rows
  payPeriods.js           -- groups dates into 2-week pay periods
  groupEntries.js          -- groups report rows by employee + period + fund
  timesheetBuilder.js       -- builds/extends the actual .xlsx (the core engine)
public/
  template/Contractor_Timesheet_PY26.xlsx  -- the blank template, bundled
                                               so new employee tabs are
                                               always built from the real
                                               thing
```

If the report's column layout changes, start in `lib/reportParser.js`. If
the timesheet template itself changes, most of the work is in
`lib/timesheetBuilder.js` (particularly the style-id constants near the
top, which were captured by hand from the current template's raw XML).

---

## Running it yourself first (optional, but recommended)

1. Install [Node.js](https://nodejs.org) (the LTS version) if you don't
   have it.
2. Open a terminal in this folder and run:
   ```
   npm install
   npm run dev
   ```
3. Open http://localhost:3000 in your browser.

---

## Deploying: GitHub → Vercel

### Part 1 — Put the code on GitHub

1. Go to [github.com](https://github.com) and click the **+** in the top
   right → **New repository**. Name it something like
   `timesheet-merger`. Leave it empty (no README/gitignore) since you
   already have those. Click **Create repository**.
2. On the next page, GitHub shows you commands under
   **"…or push an existing repository from the command line"**. Open a
   terminal in this project folder and run, one line at a time:
   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/timesheet-merger.git
   git push -u origin main
   ```
   (Replace the URL with the one GitHub showed you.) It'll ask you to sign
   in the first time — follow the prompts.
3. Refresh the GitHub page — you should see all your files there.

### Part 2 — Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) and sign in (you can sign in
   directly with your GitHub account — easiest option).
2. Click **Add New… → Project**.
3. Find `timesheet-merger` in the list of your GitHub repos and click
   **Import**.
4. Vercel auto-detects this as a Next.js app — you don't need to change
   any settings. Click **Deploy**.
5. Wait about a minute. When it finishes, click the preview thumbnail (or
   **Visit**) — that's your live URL. Bookmark it.

That's it — anyone with that link can use the tool. Every time you `git
push` a change to the `main` branch in the future, Vercel automatically
redeploys it, so updates are just: edit code → commit → push.
