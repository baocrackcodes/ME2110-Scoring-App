# ME 2110 Competition Scoring & Logistics Suite

## Master System README & Developer Guide

This suite is built on a **Single Source of Truth** architecture. The Google Spreadsheet acts as the master database, calculation engine, and manual override panel. The four web applications (Admin Scoring, Spectator Dashboard, Competitor Hub, and Broadcast Screen) act as "dumb glass"—they either feed raw data into the sheet or blindly display the pre-calculated results from the sheet.

Because of this, 90% of your competition management requires zero coding.

---

### 1. Setting Up the Master Google Sheet

Before deploying any code, your master Google Spreadsheet must be formatted exactly as the backend expects.

**Required Tabs (Exact Spelling & Capitalization):**

* **`Teams`**: Row 1 headers: `Team ID | Team Name | Sprint 3 Score`. Populate your student teams here.
* **`Scores`**: This is the raw data dump. Row 1 headers must map to the tasks in the Admin App (e.g., `Timestamp | Round | Track | Heat | Color | Section | Team ID | Team Name | Activate | Leave | Mario | Coins Inner | Coins Outer | Starbits Inner | Starbits Outer | Koopas Toppled | Koopas Displaced | Unique Lumas | Total Score | DQ? | DQ Reason`).
* **`Live_Leaderboard`**: Row 1 headers: `Rank | ID | Team Name | High Score`. (Leave the rest blank; the script overwrites it).
* **`Matches`**: Row 1 headers: `Round | Heat | Track | Yellow ID | Black ID | Blue ID | White ID | Status`. (Pre-fill your Qualifying schedule here).
* **Elimination Tabs**: Create four tabs named exactly: **`Round of 32`**, **`Sweet 16`**, **`Elite 8`**, **`Final 4`**. Row 1 headers: `Round | Heat | Track | Yellow ID | Black ID | Blue ID | White ID | Y Score | BK Score | BL Score | W Score | Status`.

*(Optional but Recommended)*: Place the master ArrayFormula in the `Scores` tab to automatically calculate and audit the final scores based on the raw task entries to prevent judge typos.

---

### 2. How to Deploy the Web Applications

You will create four separate Google Apps Script deployments.

**The Admin App (Bound to Sheet):**

1. Open your Master Google Sheet. Click **Extensions > Apps Script**.
2. Name the project "ME 2110 Admin App".
3. Paste the Admin `Code.gs` and `Index.html` files. Save.
4. Click **Deploy > New Deployment**.
5. Select type: **Web App**.
6. Execute as: **Me** (the generic ME2110 admin account).
7. Who has access: **Anyone**. (If your university requires a login, select "Anyone within [University]").
8. Click Deploy and authorize the permissions.

**The Viewing Apps (Dashboard, Competitor Hub, Broadcast):**

1. Go to `script.google.com` and click **New Project** (these do not need to be bound directly to the sheet).
2. Name the project appropriately.
3. Paste the respective `Code.gs` and `Index.html`.
4. **CRITICAL STEP:** In `Code.gs`, update the `TARGET_SHEET_ID` variable with the 44-character string from your master Google Sheet URL.
5. Deploy using the exact same steps (Execute as Me, Access: Anyone).

**⚠️ The Deployment Trap:** If you change the code later, clicking "Save" is not enough. You must click **Deploy > Manage Deployments > Edit (Pencil Icon) > Change 'Version' to 'New' > Deploy** to push the updates to the live URL.

---

### 3. Semester-to-Semester Maintenance

#### Updating Teams & Schedules

**Do not touch the code.** Simply overwrite the data in the `Teams` and `Matches` tabs of your Google Sheet. The web apps read these tabs dynamically on page load.

#### Changing the Competition Rubric (Tasks & Points)

When the game changes next semester, you must update the system in three places:

1. **The Google Sheet:** Update the column headers in the `Scores` tab to match the new tasks. Update your shadow audit array formula.
2. **The Admin UI (`Index.html`):** * Change the HTML layout to reflect the new checkboxes, inputs, and dropdowns.
* Update the `calcScore(color)` JavaScript function to multiply the new tasks by their new point values.
* Update the `payload` object in `submitScores()` to package the new task variables.


3. **The Admin Backend (`Code.gs`):**
* Update `getColMap()` to find your new column headers.
* Update `getPlayoffTiebreakerWeights()` with the new tiebreaker waterfall logic.
* Update the `rowsToWrite.push(newRow)` block inside `saveScoresToSheet()` to map the incoming payload data to the correct spreadsheet columns.



#### Editing the UI Elements

All user interfaces are built with basic HTML/CSS in the `Index.html` files.

* **Colors:** Look for the `:root` variables at the top of the `<style>` blocks (e.g., `--gt-navy`, `--gt-gold`) to change the primary theme.
* **Dropdowns:** Add or remove `<option>` tags inside the `<select>` elements in the Admin App.

---

### 4. How to Use AI for Future Development

Do not try to rewrite this architecture from scratch next semester. Use an AI (like Gemini, ChatGPT, or Claude) to update the existing framework.

**Best Practices for AI Co-Development:**

1. **Provide the Master Context:** Always paste this README and the specific `Code.gs` or `Index.html` file you want to edit into the AI prompt first.
2. **Be Explicit About the Architecture:** Remind the AI: *"Maintain the Single Source of Truth architecture. The Google Sheet handles the math; the viewing apps only display it."*
3. **Prompt Example for Rubric Updates:**
> "Act as an expert Google Apps Script developer. I am updating the ME 2110 Admin Scoring App for a new semester. Here is my current `Index.html` [paste code]. Next semester, the tasks are: 1. Pull Lever (5 pts), 2. Push Button (10 pts), 3. Spin Wheel (Drop down: Fast=20, Slow=5). Update the HTML layout, the `calcScore` function, and the `payload` object to reflect these new rules."


4. **Prompt Example for Tiebreakers:**
> "Here is my current `Code.gs` [paste code]. I need to update the `getPlayoffTiebreakerWeights()` function. The new tiebreaker hierarchy is: Tier 1 is Pull Lever points. Tier 2 is Spin Wheel points. Rewrite that specific function."



---

### 5. Troubleshooting & Critical Behaviors

* **Ghost Rows:** If the Admin App submits a score, but you don't see it at the top of the `Scores` sheet, it likely appended to row 1,000. Google Sheets `.appendRow()` and bulk injections look for the last row with *any* formatting. **Fix:** Highlight all empty rows below your data, right-click, and select "Delete Rows."
* **Manual Overrides:** To overturn a playoff result, open the relevant bracket tab (e.g., `Round of 32`), delete the ✅ emoji from the incorrect team, and paste it next to the correct team's score. The dashboards and bracket advancement script will instantly obey the change.
* **System Reset:** Use the **⚙️ ME 2110 Admin > ⚠️ RESET ALL COMPETITION DATA** custom menu in the Google Sheet to wipe all test scores and regenerate the pre-competition Sprint 3 leaderboard before the actual event begins.
