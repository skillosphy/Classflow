# Classflow

> **Your automated desktop hub for Microsoft Teams, Google Classroom & Discord.**

Classflow streamlines your academic workflow by consolidating all your courses into one clean dashboard. It automatically tracks assignments, downloads lecture slides and course materials directly to organized folders on your computer.

---

## Highlights

* **Unified 3-Column Dashboard:** View enrolled courses, active assignments, and recently downloaded lecture materials side by side.
* **Independent On-Demand & Background Sync:** Trigger assignment or lecture material syncs manually or let background automation run quietly on the interval that you set.
* **Automatic File Downloads:** Lecture slides, PDFs, documents, assignment attachments are automatically saved into neat course subfolders.
* **Google Calendar Deadline Sync:** Assignment due dates are automatically scheduled to your primary Google Calendar or any custom calendar you configure.
* **Cross-Platform Class Sources:** Seamlessly connects Microsoft Teams classes, Google Classroom courses, and Discord servers.

---

## Quick Start Guide

### 1. Download & Launch
Download **`Classflow_Setup.exe`** from the [Releases](https://github.com/skillosphy/Classflow/releases) page and install the application through the installer.

### 2. Choose Your Storage Folder
On first launch, Classflow will prompt you to select a local folder where your coursework and lecture files should be stored. Classflow automatically structures this folder by course name with dedicated `Assignments` and `Materials` subdirectories. You can change this location at any time in **Settings**.

### 3. Connect Your Classes
Open **Settings** (or click **Connect Classes** on the dashboard banner) and add your classes:
* **Microsoft Teams:** Make sure the Microsoft Teams desktop app is open, then click **Scan Teams**. Classflow will automatically discover your enrolled teams.
* **Google Classroom:** Click **Scan Classroom** to sign in securely via your web browser with your Google account and add your classrooms.
* **Discord:** Click **Add Discord**, paste your course server ID and select the channels you want to monitor for study materials or assignments or both.

### 4. Sync & Track
Click the sync icon at the top of the **Pending Assignments** or **Recently Downloaded Materials** columns for an instant refresh.

---

## Dashboard Overview

Classflow organizes your workflow into three spacious studio columns:

1. **Courses (Left Column):** Displays all your connected classes. Clicking any course card immediately opens that course's folder.
2. **Pending Assignments:** Displays active assignments ordered by due date with color coded urgency indicators. Clicking an assignment opens its attached files. 
3. **Recently Downloaded Materials:** Shows a stream of newly downloaded lecture slides, PDFs, and handouts. Clicking a material opens the file directly.

---

## Settings & Automation

* **Startup & Sync:**
  * **Run on Startup:** Launch Classflow minimized to the Windows system tray when you log in.
  * **Assignments Auto-Sync Interval:** Configurable to **6 hours**, **12 hours**, or **Daily** (default: **Daily**).
  * **Lecture Materials Auto-Sync Interval:** Configurable to **6 hours**, **12 hours**, or **Daily** (default: **Daily**).
* **Google Integration:**
  * **Google Account Authentication:** Connect or disconnect your Google Classroom and Calendar credentials.
  * **Google Calendar ID:** Sync deadlines to your default personal calendar (`primary`) or specify a dedicated calendar ID (click the **?** button for guidance on locating your Calendar ID).
* **Manage Class Connections:** Connect, disconnect or view classes across platforms at any time.

---

## Frequently Asked Questions (FAQ)

### How do I connect a Discord course?
1. In **Settings**, click **Manage** under *Connected Classes* and choose **Add Discord**.
2. Paste the Server ID of your course server.
3. If the Classflow bot is not yet present, click **Invite Bot to Server**.
4. Select the channels you want Classflow to monitor (such as #assignments, #lecture-materials) and click **Connect**.

### Does Microsoft Teams need to be open to connect?
Yes. Have the Microsoft Teams desktop app open when connecting classes so Classflow can detect your teams. 

### Where are my files stored?
All materials and assignments are saved inside your chosen download folder. Classflow creates neat subfolders for each course (e.g., `Classflow/Operating Systems/Materials` and `Classflow/Operating Systems/Assignments`).

### What happens when I close the app window?
Closing the window minimizes Classflow to the Windows System Tray. Background sync continues to run quietly on your chosen interval. Right-click the tray icon and select **Quit** to completely exit the application.

---

## Author & License

Developed by **Mohammad Hassaan**

* **GitHub:** [@skillosphy](https://github.com/skillosphy)
* **License:** Classflow is free for personal and educational use. See [LICENSE](LICENSE) for details.
