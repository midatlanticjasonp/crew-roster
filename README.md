# Crew Roster

Weekly vessel crew roster for NSML. Roster grid → Daily Breakdown printout (look-ahead and confirmed) → Excel export.
Plain HTML/JS: no build step, no server code. Data lives in **your own Firebase project** (Firestore), signed in with your Google account.

## Files

| File | What it is |
|---|---|
| `index.html`, `css/`, `js/` | The app |
| `js/firebase-config.js` | **The one file you edit** – paste your Firebase config here |
| `firestore.rules` | Database security rules – only your account can read/write |
| `firebase.json` | Firebase Hosting + rules settings |
| `starter-data.json` | Your current crew, vessels and week (22.11.26). **Not part of the repo** – load it once, then keep it private |

## Setup (about 15 minutes, once)

1. **Create a Firebase project** (or use an existing one) at <https://console.firebase.google.com>.
2. **Turn on the two services**
   - *Build → Authentication → Get started → Sign-in method → Google → Enable.*
   - *Build → Firestore Database → Create database* (production mode, nearest region e.g. `southamerica-east1` or `us-east1`).
3. **Add a web app and copy its config**: *Project settings (⚙) → Your apps → Web (`</>`) → Register app*. Copy the `firebaseConfig` object into `js/firebase-config.js` (replace the `null`).
4. **Lock the database to you**
   - Deploy once (step 5), open the app and sign in with Google. The first time it will say *"Access not allowed yet"* and show your **user ID**. (It's also under *Authentication → Users → User UID*.)
   - Paste that ID into `firestore.rules` in place of `PASTE-YOUR-USER-ID-HERE`, then publish the rules: `firebase deploy --only firestore:rules` (or paste the rules into *Firestore → Rules → Publish*).
5. **Host it** – pick one:
   - **Firebase Hosting (recommended)**: install the CLI (`npm i -g firebase-tools`), then in this folder:
     `firebase login` → `firebase use --add` (pick your project) → `firebase deploy`.
     Your app is at `https://<project-id>.web.app`.
   - **GitHub Pages**: push this folder to a repo, *Settings → Pages → Deploy from branch*. Then add the Pages domain
     (e.g. `yourname.github.io`) in Firebase *Authentication → Settings → Authorized domains*, or Google sign-in will be refused.
6. **Load your data**: in the app go to *Settings & Data → Restore from backup…* and pick `starter-data.json`.
   Afterwards delete that file from anywhere shared – it contains crew phone numbers and emails.

> **Never commit crew data to GitHub.** `.gitignore` already excludes `starter-data*.json` and backup files. If your repo is public,
> that's fine for the code (the Firebase config is not a secret – the rules are what protect the data).

**Trying it before Firebase is set up:** leave `firebaseConfig = null` and serve the folder locally (`npx serve .` or
`python -m http.server`), then open the address it prints. A yellow bar shows you're in *local test mode* – data stays in that browser only.
(Double-clicking `index.html` won't work; browsers block the app's modules from `file://`.)

## Weekly routine

1. **Monday – Start new week.** Everyone carries over pre-filled with their last code for all 7 days; anyone whose last entry was **OFF** is removed. People on leave/sick/absent carry over *with their note*. Last week stays saved – reopen it from the week list.
2. **Change only the exceptions** on the Roster tab:
   - Pick a code from the dropdown. On an empty line, the first code you pick fills the rest of the week.
   - **Crew change:** leaver → **OFF** + time (later days clear automatically). Joiner → **+ Add line**, pick position and name, **ON** + time on the joining day (later days fill with W).
     Set the joiner's **Relieving** to the leaver to print both in one cell.
   - Leave / sick / absent / came off early: just change the code and add a **Note** – the person keeps their line on the printout. No line counts to maintain.
3. **Print look-ahead** → *Save as PDF* (file name is pre-set: `Crew Roster WE dd.mm.yy - LOOK-AHEAD`).
4. During the week, update as things happen.
5. **Checks** tab: fix every red **FIX** item.
6. **Print confirmed** → *Save as PDF* for payroll. If anything changes afterwards the status turns to *Changed after confirm* so you know to re-send.

**Backups:** *Settings & Data → Download backup* saves everything (all weeks) as one `.json` file on your computer – do it monthly. *Restore from backup* loads it back.

## Codes

Edit in *Settings & Data → Day codes*: label printed on the Daily Breakdown, whether it counts toward manning, whether it's paid, whether a note is required.
W / ON / OFF are built in.

## How the pieces fit (for future changes)

- `js/model.js` – all roster rules: printout layout, checks, new-week carry-over. No screen code.
- `js/app.js` – screens and buttons.
- `js/export.js` – Excel file (ExcelJS, loaded from cdnjs).
- `js/store.js` – saving: Firestore (`app/meta`, `app/crew`, `weeks/<week-ending>`) with offline cache, or browser storage in test mode.
