# 🔥 Rising Phoenix

A powerful Google Sheets + Apps Script productivity system for tracking study progress, projects, habits, fitness, and personal growth.

**Rising Phoenix** combines a customizable daily checklist, streak tracking, progress analytics, motivational dashboard, and automated reporting into a single self-maintaining workspace. Built around flexible segments and tasks that can be edited at any time without losing your history. 【1-8128eb】

---

# ✨ Features

## 🏠 Home Dashboard

A dedicated dashboard designed as the main landing page of the workbook.

Features:

- Daily motivational quotes
- Current streak counter
- Weekly average score
- Today's completion score
- Quick statistics
- Segment score tracking
- Configurable streak target
- Live dashboard formulas
- Blue-themed interface

No more empty startup sheet. The first tab becomes a fully functional dashboard. 【1-8128eb】

---

## ✅ Dynamic Daily Checklist

Track anything you want:

- Study sessions
- University courses
- Projects
- Coding practice
- Creative work
- Fitness
- Personal growth

Features:

- Automatic daily columns
- Checkbox tracking
- Comments per task
- Automatic completion percentages
- Unlimited historical tracking

Unlike previous versions, nothing is hardcoded. Segments and tasks are fully configurable. 【1-8128eb】

---

## ⚙️ Configurable Segments and Tasks

Open:

```text
Dashboard
└── Edit segments & tasks...
```

You can:

- Rename segments
- Add new segments
- Remove segments
- Reorder segments
- Rename tasks
- Add tasks
- Remove tasks

Historical checkbox data and comments are preserved whenever possible. 【1-8128eb】

---

## 📈 Weekly Progress

Automatically tracks:

- Segment-by-segment completion
- Daily progress trends
- Weekly averages
- Performance over time

Includes an automatically generated line chart with weekly average overlay. 【1-8128eb】

---

## 📊 Daily Completion

Tracks:

- Overall completion percentage
- Daily consistency
- Streak calculations
- Completion history

The home dashboard pulls its statistics directly from this sheet. 【1-8128eb】

---

## 🔥 Streak System

Built-in streak tracking system.

Default behavior:

```text
70% completion or higher
=
Successful Day
```

Features:

- Configurable target
- Current streak tracking
- Dashboard integration
- Automatic calculation

---

## 🎨 Visual Feedback

Automatic:

- Checkbox generation
- Percentage calculations
- Conditional formatting
- Gradient completion tracking
- Daily completion charts
- Weekly progress charts

Completion percentages are colorized for instant visual feedback. 【1-8128eb】【2-8513d8】

---

## 🤖 Full Automation

The system automatically:

- Adds new daily columns
- Updates progress values
- Refreshes charts
- Updates dashboard calculations
- Changes motivational quotes
- Maintains statistics

A daily trigger runs automatically at:

```text
6:00 AM
```

No manual maintenance required. 【3-5b5b99】【4-da285f】

---

# 🏗 Workbook Structure

```text
🔥 Rising Phoenix
│
├── Home
│   ├── Motivation
│   ├── Current Streak
│   ├── Weekly Average
│   ├── Daily Completion
│   ├── Quick Stats
│   └── Segment Scores
│
├── Daily Checklist
│
├── Weekly Progress
│
├── Daily Completion
│
└── Chart Data
    (auto-generated helper sheet)
```

【1-8128eb】

---

# 📸 Screenshots

## Home Dashboard

- Motivation
- Statistics
- Streak Tracking
- Segment Scores
<img width="1920" height="1048" alt="homepage" src="https://github.com/user-attachments/assets/b3e00210-79c2-4924-8756-2fa0531502f6" />

<img width="1920" height="1048" alt="chartdata" src="https://github.com/user-attachments/assets/657e88af-0ae6-47c2-94a7-c6c1b9c69e6e" />
<img width="1920" height="1048" alt="dashboard" src="https://github.com/user-attachments/assets/7d5df107-fa82-46e8-beeb-7154282fe7cf" />

---

## Daily Checklist

<img width="1920" height="1048" alt="daily_checklist" src="https://github.com/user-attachments/assets/991b1e7f-8994-4c2f-9fe4-a9a84bb51369" />


Features:

- Segments
- Tasks
- Comments
- Checkboxes
- Completion %

---

## Weekly Progress
<img width="1920" height="1048" alt="weeklyprogress" src="https://github.com/user-attachments/assets/04147783-c34c-4085-94b7-8b47b5e24d35" />

Features:

- Trend chart
- Weekly averages
- Segment comparison

---

## Daily Completion

<img width="1920" height="1048" alt="daily_completition" src="https://github.com/user-attachments/assets/22f179e7-6d82-431d-bfaa-1fae75b29651" />


Features:

- Overall completion percentage
- Daily consistency tracking

---

# ⚙️ Installation

## 1. Create a Google Spreadsheet

Create a new spreadsheet.

---

## 2. Open Apps Script

```text
Extensions
└── Apps Script
```

---

## 3. Paste the Code

Copy the complete contents of:

```text
Code.gs
```

into the Apps Script editor.

---

## 4. Run Setup

Execute:

```javascript
setupDashboard()
```

Authorize the requested Google permissions.

---

## 5. Done

Setup automatically:

- Creates sheets
- Creates formulas
- Builds dashboard
- Creates charts
- Installs triggers
- Configures formatting

After setup, the workbook becomes fully automated. 【1-8128eb】

---

# 📋 Dashboard Menu

After installation a custom menu appears:

```text
Dashboard
├── Edit segments & tasks...
├── Add today's column now
├── Refresh charts
├── New motivational quote
├── Rebuild home page
├── Run setup (erases history)
├── Seed 7 days of test data
└── Reset names to defaults
```

【1-8128eb】

---

# 🧪 Testing

Generate demo data:

```javascript
seedRandomWeekData()
```

Creates:

- 7 days of history
- Realistic completion data
- Dashboard statistics
- Chart previews

Perfect for testing before real use. 【1-8128eb】

---

# 🔄 Typical Daily Workflow

1. Open **Rising Phoenix**
2. Review today's dashboard
3. Check the motivational quote
4. Complete checklist items
5. Monitor progress in real time
6. Return tomorrow for a fresh day automatically

No manual column creation is required.

---

# ⚠ Important Notes

## Initial Setup

Run:

```javascript
setupDashboard()
```

only once during installation.

---

## Changing Layout

To modify segments or tasks:

```text
Dashboard
└── Edit segments & tasks...
```

Avoid manually inserting or deleting rows inside the checklist sheet. 【1-8128eb】

---

## Triggers

Verify the trigger exists:

```text
Apps Script
└── Triggers
```

You should see:

```text
Function: addDailyColumn
Frequency: Daily
Time: 6:00 AM
```

【3-5b5b99】【4-da285f】

---

# 🛠 Main Functions

| Function | Purpose |
|-----------|-----------|
| `setupDashboard()` | Initial installation |
| `addDailyColumn()` | Add a day manually |
| `refreshCharts()` | Regenerate charts |
| `rebuildHome()` | Rebuild dashboard |
| `newQuote()` | Generate a new quote |
| `seedRandomWeekData()` | Create demo data |
| `resetToDefaults()` | Restore defaults |
| `createTestTrigger()` | Testing automation |
| `deleteAllTriggers()` | Remove all triggers |

【1-8128eb】

---

# 📄 License

MIT License

Free to use, modify, distribute, and build upon.

---

# 🤝 Contributing

Contributions are welcome.

Possible future improvements:

- Best streak tracking
- Achievement levels
- Semester templates
- Exportable reports
- Mobile-friendly layouts
- Advanced analytics
- Focus recommendations

Feel free to open an issue or submit a pull request.

---

## Built With

- Google Sheets
- Google Apps Script
- Custom Dashboard UI
- Automated Charts
- Dynamic Configuration System
- Streak Analytics
- Productivity Tracking Framework
- helps of AI

**Rise. Check. Repeat. 🔥**
