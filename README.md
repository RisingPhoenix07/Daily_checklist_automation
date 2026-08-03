# Daily Checklist Dashboard

A Google Sheets + Apps Script automation for tracking study progress, daily tasks, and weekly averages — complete with gradient formatting and auto‑generated charts.

---

## 🚀 Features
- Automated **Daily Checklist** with segments, tasks, and % Complete rows.
- **Weekly Progress** sheet with segment tracking + Weekly Average overlay.
- **Daily Completion** sheet logging overall % per day.
- Gradient formatting for instant visual feedback.
- Charts with **percentage axes** (Weekly line chart + Daily column chart).
- Automatic **daily trigger** at 6 AM to add new columns hands‑free.

---

## 📸 Screenshots

<img width="1920" height="1048" alt="Screenshot From 2026-08-03 18-32-46" src="https://github.com/user-attachments/assets/9648bee6-b49a-423b-93bd-c02cb102cfd1" />
<img width="1920" height="1048" alt="Screenshot From 2026-08-03 18-32-41" src="https://github.com/user-attachments/assets/8184803f-7f1b-472e-9bd1-32b0bb074fad" />
<img width="1920" height="1048" alt="Screenshot From 2026-08-03 18-32-30" src="https://github.com/user-attachments/assets/329a4aa7-191c-41be-bc06-1ff87a45d98a" />

---

## ⚙️ Setup
1. Open your Google Sheet → Extensions → Apps Script.
2. Paste the full script from this repo into `Code.gs`.
3. Run `setupDashboard()` once to initialize everything.
4. Authorize the script when prompted.
5. Done — your dashboard is live and automated.

---

> **Important Notice**  
> For normal, ongoing use, you only need to run `setupDashboard()` once.  
> That single run:  
> - Builds all three sheets with your segment/task names  
> - Sets up the gradient formatting and both charts  
> - Installs the 6 AM daily trigger  
>
> After that, everything is automatic: the trigger fires `addDailyColumn()` every day at 6 AM on its own. You just open the sheet and check off boxes throughout the day.  
>
> You’d only manually run something again if:  
> 1. You change `SEGMENT_NAMES` or `TASK_NAMES` in the script → re‑run `setupDashboard()` to rebuild the checklist with the new layout (this wipes existing data, so only do it when you want a reset).  
> 2. You want to add test/demo data → run `seedRandomWeekData()`.  
> 3. Triggers get into a weird state → run `deleteAllTriggers()`, then `setupDashboard()` again to reinstall a clean one.  
>
> ### One thing worth checking once  
> Since you’re switching from test data to real daily use: open **Triggers** in the left sidebar of the Apps Script editor (clock icon) and confirm there’s exactly one trigger listed for `addDailyColumn`, running daily at 6 AM. That confirms `setupDashboard()` installed it correctly.

---

## 🧩 Usage
- **Daily automation**: Trigger runs `addDailyColumn()` every morning.
- **Manual refresh**: Run `addDailyColumn()` manually if needed.
- **Testing**: Use `createTestTrigger()` for minute‑by‑minute testing, then `deleteAllTriggers()` to clean up.

---

## 🛡 License
MIT License — free to use, modify, and share.

---

## 🤝 Contributing
Pull requests welcome! For major changes, open an issue first to discuss what you’d like to change.
