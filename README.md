# BHAVE

Fast behavior logging for classroom aides. Runs from GitHub Pages and saves to a Google Sheet — no Claude account needed.

## How aides use it
1. **Tap a student.**
2. Tap **+** (green) or **−** (red) next to a behavior. The switch at the bottom decides what each tap means:
   - **Count** — each **+** is one occurrence.
   - **Time (minutes)** — each **+** is one minute (tap in how long it lasted).
3. Tap **Submit**. **Cancel** throws it away.

The first time the app opens it asks who is using the phone; that name is recorded with every entry.

## For the teacher
- **Students and behaviors:** edit the **Setup** tab of the Google Sheet. One row per behavior: `Initials | Behavior | Color (optional, e.g. #1E88E5)`. Changes show up on every phone the next time BHAVE opens.
- **Staff names:** edit the **Staff** tab, one name per row (it starts with Para A, Para B, Teacher). These are the "Who's using this phone?" choices. To change the name on a phone, tap the name chip at the top of the student screen.
- **Data:** every entry is a row in the **Log** tab: Date, Time, Student, Behavior, Amount, Type (count or minutes), Logged by.
- **Script:** `apps-script/Code.gs` is the Google Apps Script that runs inside the Sheet (this copy is for your records).
- Use initials only, never full names.
