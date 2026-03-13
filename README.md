# hem_home-expenses-mobile

Mobile web app for entering expenses into the **Home Expenses YYYY** Google Sheets spreadsheet. Hosted as a Google Apps Script web app and embeddable in Google Sites.

## Features

- **One-Time (OT)** — adds expense to the current month only
- **Recurrent (RM)** — adds expense from the current month through December
- Category dropdown populated dynamically from the spreadsheet and accepts new free-text values
- Same fields as the desktop form: Name, Category, Amount, Split, PAP, Paid, Billing Period, Who Paid
- **Settings page** (`?page=settings`) — configure the target spreadsheet by URL, ID, or Drive file name
- **Auto year-switch** — optional Jan 1 trigger that finds the new *Home Expenses YYYY* file automatically

## Project Structure

```
appsscript.json          GAS manifest (webapp, ANYONE_WITH_GOOGLE_ACCOUNT)
.clasp.json              clasp config (fill in Script ID)
src/
  StaticNumbers.js       Column/row constants (keep in sync with HomeExpenses project)
  MobileAddExpense.js    doGet(), mobileProcessForm(), settings functions, year-switch trigger
ui/
  MobileAddExpense.html  Mobile expense form (OT/RM tabs, ⚙️ settings link)
  Settings.html          Settings page — spreadsheet config + auto year-switch toggle
```

## One-Time Setup

### 1. Create a new Google Apps Script project

Go to [script.google.com](https://script.google.com), create a new project, and copy its **Script ID** from Project Settings.

### 2. Update `.clasp.json`

Replace `YOUR_SCRIPT_ID_HERE` with the real Script ID.

### 3. Push code

```bash
cd /home/alexs/projects/hem_home-expenses-mobile
clasp push
```

### 4. Deploy as a Web App

In the Apps Script editor:

1. **Deploy → New deployment**
2. Type: **Web app**
3. Execute as: **Me**
4. Who has access: **Anyone with Google Account**
5. Click **Deploy** and copy the web app URL

### 5. Configure the spreadsheet via the Settings page

#### Navigating to Settings

Append `?page=settings` to your web app URL:

```
https://script.google.com/macros/s/<deployment-id>/exec?page=settings
```

You can also tap the **⚙️** gear icon inside the expense form at any time.

---

#### Connected Spreadsheet panel

At the top of the Settings page, a badge shows the current connection state:

| Badge | Meaning |
|-------|---------|
| 📊 *file name* — "Currently connected" | A spreadsheet is configured and accessible |
| ⚠️ "Not configured — Enter a URL or file name below" | No spreadsheet has been set yet |
| 📊 "(file not accessible — ID may be stale)" | The saved ID is no longer reachable (file moved, deleted, or permission revoked) |

---

#### Entering the spreadsheet

In the **Spreadsheet URL, ID, or Drive File Name** field, enter **any one** of the following:

1. **Full Google Sheets URL** — paste the URL directly from your browser's address bar:
   ```
   https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/edit
   ```
   The script extracts the ID automatically from the `/d/<id>/` segment.

2. **Spreadsheet ID only** — the 44-character alphanumeric string found between `/d/` and `/edit` in the URL:
   ```
   1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms
   ```

3. **Exact Drive file name** — the app searches your Google Drive for a file whose name matches exactly (case-sensitive):
   ```
   Home Expenses 2026
   ```
   > **Note:** The file must be owned by or explicitly shared with the Google Account that the web app runs as ("Execute as: Me"). If Google Drive returns no result, double-check the name and sharing settings.

Click **Connect Spreadsheet**. A loading spinner appears while the app validates the input.

---

#### What happens on save

The `saveSpreadsheetConfig` function on the server:

1. Tries to parse the input as a URL or raw ID first (`/d/<id>/` pattern or 25+ alphanumeric chars).
2. If that fails, performs a Drive filename search (`DriveApp.getFilesByName`).
3. On success, stores the resolved spreadsheet ID in **Script Properties** under the key `SPREADSHEET_ID` and returns the file's display name.
4. The badge on the page updates immediately to 📊 *file name* — "Currently connected" without a page reload.

---

#### Error messages

| Message | Cause & fix |
|---------|-------------|
| *"Please enter a URL or file name."* | The input field was empty. |
| *"Could not open spreadsheet. Check the URL/ID…"* | The ID was parsed but `SpreadsheetApp.openById` failed — the file may be in a different account, deleted, or not shared. |
| *"File \"…\" not found in Google Drive."* | Drive name search returned nothing — verify the exact name and that the file is in the same Google account. |

---

#### Verifying the connection

After a successful save, go back to the main form (tap **← Back to form** or remove `?page=settings` from the URL). The form header will show the connected spreadsheet name, and the category dropdown will populate from live spreadsheet data.

### 6. Enable Auto Year-Switch (optional)

On the Settings page, toggle **Switch file on January 1st** on.  
Each New Year's Day at 6 AM the script will search your Drive for `Home Expenses {year}` and switch automatically. No action needed when you copy the spreadsheet to the new year.

### 7. (Optional) Embed in Google Sites

In Google Sites, add an **Embed** block with the web app URL.

## Keeping Column Constants in Sync

`src/StaticNumbers.js` is a trimmed copy of `staticNumbers` from `HomeExpenses/src/static/main.js`.  
If column positions change in the main project, update this file to match.
