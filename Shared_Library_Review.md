# Sharing `StaticNumbers` between Projects

Yes, you can absolutely share the `StaticNumbers` class between your main `HomeExpenses` project and your `hem_home-expenses-mobile` project without duplicating code.

Here are the two best approaches depending on your workflow:

## Option 1: Google Apps Script Libraries (The Native GAS Way)
Google Apps Script has a built-in feature specifically for sharing code between projects called **Libraries**.

### How to do it:
1. **Deploy the Main Project**: In your main `HomeExpenses` project (in the online GAS editor), click **Deploy > New deployment**, choose **Library**, and deploy it.
2. **Get the Script ID**: Go to **Project Settings** (the gear icon) in the main project and copy the **Script ID**.
3. **Add to Mobile App**: In your `hem_home-expenses-mobile` online editor, click the **"+"** next to **Libraries** on the left sidebar. Paste the Script ID, select the version you just deployed, and give it an identifier (e.g., `HE`).
4. **Usage**: In your mobile app code, you would access it like this:
   ```javascript
   var myNumbers = new HE.staticNumbers();
   ```

* **Pros:** It's built exactly for this purpose and keeps things cleanly separated.
* **Cons:** Every time you change a number in the main repo, you have to create a new deployment version of the library, and then update the mobile app to point to that new version.

---

## Option 2: Local File Sync / Symlink (The Clasp / Developer Way)
Since you are working locally with both repositories on your Linux machine (likely using `clasp` to push code), you can just share the physical file locally. When you `clasp push`, Google will just see it as a normal file in both projects.

### How to do it:
You can create a symbolic link (symlink) in your mobile project that points to the file in your main project.
Run this in your terminal:
```bash
# Delete the duplicate file in the mobile repo
rm /home/alexs/projects/hem_home-expenses-mobile/src/static/StaticNumbers.js

# Create a symlink pointing to the main repo's file
ln -s /home/alexs/projects/HomeExpenses/src/static/mail.js /home/alexs/projects/hem_home-expenses-mobile/src/static/StaticNumbers.js
```
*(Note: adjust the path to your HomeExpenses repo if it's different!)*

* **Pros:** 
  - **Zero runtime overhead:** Google Apps Script just sees it as a native file in both projects, so it runs at maximum speed.
  - **No versioning headaches:** If you update the file in your main repo, the change is instantly reflected in your mobile repo the next time you run `clasp push`. You don't have to manage library versions.

## Recommendation
* If you want to manage everything via the web editor, go with **Option 1 (Libraries)**. 
* If you prefer doing your development locally using `clasp`, **Option 2 (Symlink)** is far superior because you never have to worry about library versions falling out of sync.
