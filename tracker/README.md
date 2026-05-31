# Daily IT Project Manager Excel Tracker

A Python script that generates a fully formatted, multi-sheet Excel workbook (`daily_project_tracker.xlsx`) designed for an experienced IT Project Manager to manage daily development tasks across multiple teams.

## What the Tracker Does

The generated workbook contains **7 sheets** covering all aspects of a typical IT project sprint:

| Sheet | Purpose |
|---|---|
| **DASHBOARD** | High-level sprint summary with live metrics (ticket counts by status, open issues) |
| **SPRINT TRACKER** | Core sprint board — all tickets with status, assignee, priority, blockers, and traceability flags |
| **PRODUCTION ISSUES** | Log of active and resolved PROD incidents with severity, RCA, and comms tracking |
| **UAT TRACKER** | UAT handoff and sign-off tracking per feature/story |
| **QA ISSUES** | QA bug log with retest tracking and UAT/PROD blocker flags |
| **PROD RELEASE TRACEABILITY** | End-to-end release matrix linking Dev → QA → UAT → PROD deploy |
| **DAILY STANDUP LOG** | Daily standup notes per team member with blocker and action tracking |

## Installation

```bash
pip install -r requirements.txt
```

## How to Run

```bash
python generate_tracker.py
```

This produces `daily_project_tracker.xlsx` in the current working directory.

## Sheet Descriptions

### DASHBOARD
Provides a sprint header (name, start/end dates) and a summary metrics table with `COUNTIF`-style formulas linked to the Sprint Tracker sheet. Status counts (Done, In Progress, Blocked, To Do) and open issue counts for PROD, UAT, and QA are displayed with color-coded cells.

### SPRINT TRACKER
The main sprint board with columns for Ticket ID, Summary, Team, Assignee, Priority, Status, Story Points, dates, blocker notes, and traceability flags (Going to QA/UAT/PROD). Supports 5 sample rows across Frontend, Backend, and QA teams. Conditional formatting highlights Blocked tickets in red, Done in green, and In Progress in amber.

### PRODUCTION ISSUES
Tracks live and historical PROD incidents including severity (P1–P3), environment, root cause, fix ticket, and whether a postmortem and comms were completed. P1 issues are highlighted red, P2 orange.

### UAT TRACKER
Tracks each feature sent to UAT — owner, status (Pending / In Testing / Passed / Failed / Waived), defects raised, sign-off details, and PROD readiness. Passed rows are green, Failed red, In Testing amber.

### QA ISSUES
Bug log from QA testing. Tracks severity, assigned developer, retest dates and results, and flags whether the issue blocks UAT or PROD. Open issues are red, Closed are green.

### PROD RELEASE TRACEABILITY
The release governance matrix. Each row represents a release item traced from its sprint ticket through QA, UAT, and into PROD deployment. Tracks deploy status, who deployed, post-deploy verification, and rollback plan.

### DAILY STANDUP LOG
Rolling log of standup updates per team member — yesterday, today, blockers, action owners, and resolution status.
