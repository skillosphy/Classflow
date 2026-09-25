# Classflow

> **Your automated desktop hub for Microsoft Teams, Google Classroom & Discord.**

Classflow streamlines your academic workflow by consolidating all your courses into one clean, modern dashboard. It automatically tracks assignment deadlines, downloads lecture slides and course materials directly to organized folders on your system.

---

## Highlights

* **Unified 3-Column Dashboard:** View connected courses, pending assignments, and course materials side by side on the dashboard.
* **Independent Background & On-Demand Sync:** Refresh assignments or course materials instantly with manual sync, or let background automation run quietly on intervals you configure.
* **Automatic File Downloads:** Lecture slides, PDFs, code files, and assignment attachments are automatically saved into structured course subfolders.
* **Cross-Platform Integration:** Seamlessly connects courses across Microsoft Teams, Google Classroom, and Discord servers.
* **Instant File Access:** Click any assignment or material to open the document directly where it has been saved.

---

## Quick Start Guide

### 1. Download & Launch
Download the installer from the [Releases](https://github.com/skillosphy/Classflow/releases) page and run.

### 2. Choose Your Download Folder
On first launch, Classflow will prompt you to select a folder on your computer where course materials and assignments should be saved. Classflow automatically organizes files into neat course subdirectories. You can change this location at any time in **Settings**.

### 3. Connect Your Classes
* **Microsoft Teams:** Ensure the Microsoft Teams desktop app is open, then click **Scan Teams**. Classflow will automatically discover your enrolled teams and SharePoint class sites.
* **Google Classroom:** Click **Scan Classroom** to sign in securely via your browser with your Google account and select your courses.
* **Discord:** Click **Add Discord**, paste your course server invite link/any channel link, and select which channels you want Classflow to monitor (for materials, assignments, or both).

### 4. Sync & Track
Click the sync icon at the top right of the dashboard, or let background auto sync handle everything quietly.

---

## Dashboard Overview

Classflow organizes your workflow into three columns:

1. **Courses:** Displays all your connected classes across Teams, Classroom, and Discord with color-coded bars. Clicking a course card immediately opens that course's folder on your system.
2. **Pending Assignments:** Displays upcoming active assignments ordered by due date with color-coded urgency indicators (green for 3+ days, yellow for 1–2 days, red for today). Clicking an assignment card opens its attached files directly.
3. **Course Materials:** Streams newly downloaded lecture slides, readings, and handouts. The feed is sectioned into **Today** and **Previous Downloads**, accompanied by a count `(X new)` of newly downloaded materials. Clicking any material opens it immediately.

---

## Settings & Automation

* **Startup & Sync:**
  * **Run on Startup:** Enabled by default. Automatically adds Classflow to start when you boot your system. Can be toggled on/off at any time.
  * **Assignments Auto-Sync Interval:** Configurable to **6 hours**, **12 hours**, or **Daily** (default: **Daily**).
  * **Course Materials Auto-Sync Interval:** Configurable to **6 hours**, **12 hours**, or **Daily** (default: **Daily**).
* **Connected Classes:**
  * **Manage Class Connections:** View, configure, or disconnect classes across platforms at any time.

---

## Frequently Asked Questions (FAQ)

### Does Microsoft Teams need to be open to connect?
Yes, have the Microsoft Teams desktop app open when initially scanning and connecting classes so Classflow can capture your active school credentials. Subsequent syncs can run headlessly via cached credentials.

### Where are my files stored?
All materials and assignments are saved directly inside your configured download directory. Classflow automatically manages subdirectories for each course:
* `<Download Folder>/<Course Name>/Materials/`
* `<Download Folder>/<Course Name>/Assignments/`

### What happens when I close the app window?
Closing the window minimizes Classflow to the Windows System Tray. Background auto-sync continues to run on your chosen schedule. Right-click the tray icon and click **Quit** to close the application completely.

### Why does Windows SmartScreen show a warning?
Classflow is an independent educational tool without an expensive corporate code-signing certificate. Windows SmartScreen may show a *"Windows protected your PC"* prompt on first launch.
To proceed:
1. Click **More info**.
2. Click **Run anyway**.

### Why does Google say "Google hasn't verified this app" during Google Auth?
When authenticating your Google account to sync Google Classroom classes, Google may display a prompt stating *"Google hasn't verified this app"*. This is standard for apps that have not gone through Google's review program.
To continue:
1. Click **Advanced** (or *Show Advanced*) at the bottom-left of the Google auth prompt.
2. Click **Go to Classflow (unsafe)** to grant access.

---

## Author & License

Developed by **Mohammad Hassaan**

* **GitHub:** [@skillosphy](https://github.com/skillosphy)
* **License:** Classflow is free for personal and educational use. See [LICENSE](LICENSE) for details.
