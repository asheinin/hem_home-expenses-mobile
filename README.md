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

Open `<your-web-app-url>?page=settings` and enter either:
- The full Google Sheets URL of *Home Expenses 2026*
- The spreadsheet ID (from the URL)
- The exact Drive file name, e.g. `Home Expenses 2026`

Click **Connect Spreadsheet** — the app stores the ID as a Script Property.

### 6. Enable Auto Year-Switch (optional)

On the Settings page, toggle **Switch file on January 1st** on.  
Each New Year's Day at 6 AM the script will search your Drive for `Home Expenses {year}` and switch automatically. No action needed when you copy the spreadsheet to the new year.

### 7. (Optional) Embed in Google Sites

In Google Sites, add an **Embed** block with the web app URL.

## Keeping Column Constants in Sync

`src/StaticNumbers.js` is a trimmed copy of `staticNumbers` from `HomeExpenses/src/static/main.js`.  
If column positions change in the main project, update this file to match.
