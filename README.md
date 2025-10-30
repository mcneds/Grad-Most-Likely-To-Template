# Most Likely To — Template

Anonymous “most likely to…” nominations with per‑device spam control, up‑voting, and simple moderation.
Frontend runs on **GitHub Pages** (static). Backend is a **Google Apps Script** bound to a Google Sheet.

**Highlights**

* Multiple submissions per device ✅
* A device can’t nominate the **same person twice** ✅
* Anonymous responses (device-only token) ✅
* One like per device per response ✅
* Admin deletion & admin bypass ✅
* Free hosting (GitHub Pages + Apps Script) ✅

---

## Contents

* [Quick start (10–15 min)](#quick-start-1015-min)
* [Google Sheet schema](#google-sheet-schema)
* [Admin actions](#admin-actions)
* [Customization](#customization)
* [Data & exports](#data--exports)
* [Troubleshooting](#troubleshooting)
* [FAQ](#faq)
* [Repo structure](#repo-structure)
* [License](#license)

---

## Quick start (10–15 min)

### 0) Create (or copy) the Google Sheet

Make a Sheet with **four tabs** (see schema below). You can also copy an existing sheet.

### 1) Add names

On the **Names** tab, fill two columns:

```
A: id    B: name
1        Alex Kim
2        Priya Singh
3        …
```

> Tip: Paste a single column of names into **B**, then drag‑fill numbers in **A**.

### 2) Backend — Apps Script

1. Open your Sheet → **Extensions → Apps Script**.
2. Replace everything in `Code.gs` with the code from `/backend/Code.gs` (this repo).
3. **Project Settings → Script properties** → add:

   * `ADMIN_KEY` → any secret string (e.g., `MLT_2026_Sec!ret`).
4. **Deploy → New deployment → Web app**

   * Execute as: **Me**
   * Who has access: **Anyone**
   * Copy the **/exec URL** (you’ll paste it in the frontend).

> The script is **container‑bound** to the sheet (`SpreadsheetApp.getActive()`), so it must live inside the target Sheet.

### 3) Frontend — get `index.html` from `/frontend/`

If you’re using this repo as a template, the latest `index.html` lives in **`/frontend/index.html`**.

* Option A (recommended): **Copy** `/frontend/index.html` to your new repo’s root as `/index.html`.
* Option B: Keep the file in `/frontend/` and add a tiny root redirect file:

  ```html
  <!-- /index.html at repo root -->
  <meta http-equiv="refresh" content="0; url=frontend/index.html">
  <link rel="canonical" href="frontend/index.html">
  ```

Then open `/index.html` and set the config at the top:

```js
const CONFIG = {
  SCHOOL_NAME: "Your School Name",
  WEB_APP_URL: "https://script.google.com/macros/s/…/exec", // paste your Apps Script web app URL
  USE_SOFT_FINGERPRINT: true,
  VERSION: "v0.1.x"
};
```

Commit the change.

### 4) Turn on GitHub Pages in a **new repo**

If you’re starting fresh:

1. On GitHub, click **Use this template** (or **Fork**) to create your own repo.
2. Add `/index.html` at the **repo root** (copy from `/frontend/index.html` as above).
3. Go to **Settings → Pages**:

   * **Source**: *Deploy from a branch*
   * **Branch**: your default branch (e.g., `main`)
   * **Folder**: **/** (root)
4. Click **Save**. GitHub will publish your site at:
   `https://<your-username>.github.io/<repo-name>/`

> If you prefer, you can also use the `/docs` folder: move `index.html` into `/docs/` and choose **/docs** in Pages settings.

---

## Google Sheet schema

Create these tabs with **exact** headers in Row 1.

### Names

```
A: id     B: name
1         Alex Kim
2         Priya Singh
…
```

### Nominations *(written by script)*

```
A: timestamp    B: device_id               C: name_id   D: text
2025‑10‑30      6d3b…#-123456789            2           most likely to…
```

### Votes *(written by script)*

```
A: timestamp    B: device_id               C: nomination_id
2025‑10‑30      6d3b…#-123456789            17
```

### Settings *(optional)*

Not required by the code (the admin key uses Script Properties). You can keep other notes/values here.

> `device_id` is a random per‑browser value (with a soft fingerprint after `#`). It is not personally identifying.

---

## Admin actions

* **Enter admin mode** (hidden):

  * **Ctrl + Alt + A** (works only when not focused in a field), **or**
  * **Click the page title** (“Most Likely To”) **5 times quickly**.
* Enter your `ADMIN_KEY` when prompted.
* In **Responses** view, a **Delete** button appears under each nomination.
* Admins can also submit even if the same device already nominated the same person (bypass rule).

---

## Customization

* **Branding**: update the header text and `SCHOOL_NAME` in `index.html`.
* **Version tag**: bump `CONFIG.VERSION` to confirm deploys (shows in the footer, and logs in the console).
* **New Submission**: the top‑right button returns to the form without clearing device data.
* **Spam control**:

  * Multiple submissions **allowed** per device.
  * Duplicate **(device_id, name_id)** nominations are blocked **server‑side**.
  * Voting is one like per device per nomination.

---

## Data & exports

* All data lives in your Google Sheet.
* Export via **File → Download** (CSV/Excel) or build charts/pivots directly in the Sheet.
* “Nomination ID” equals the **row number** in the Nominations tab (stable ID used for votes & deletes).

---

## Troubleshooting

**404 on Pages**

* Ensure an `index.html` exists at the **repo root** (Pages serves `/` or `/docs`).
* If you keep `/frontend/index.html`, create a root redirect:

  ```html
  <meta http-equiv="refresh" content="0; url=frontend/index.html">
  ```

**CORS / “Failed to fetch”**

* Frontend `fetch()` uses **no custom headers** (to avoid preflight). Keep it that way.
* In Apps Script, re‑**Deploy → New deployment** after code changes and use the **new `/exec` URL**.
* **Access** must be set to **Anyone**.

**Frontend won’t update**

* Bump `CONFIG.VERSION` and **hard refresh** (`Ctrl/Cmd+Shift+R`). GitHub Pages may cache.

**Admin prompt appears while typing**

* You’re on an old build that used the `A` key. Current build only uses **Ctrl+Alt+A** and the **5‑click title** trigger.

**No names appear**

* Check the **Names** headers: exactly `id | name`. Ensure rows have values.

**Votes/Nominations not recording**

* Verify the tab names and header text match exactly. Apps Script is case‑sensitive.

---

## FAQ

**Are submissions anonymous?**
Yes. The app stores a random per‑browser `device_id` (with a soft fingerprint). No login or personal info is collected.

**Can students submit multiple times?**
Yes, but a device cannot submit for the **same person twice**.

**Can I seed responses or fix typos?**
Yes. Enter admin mode and submit (admins bypass duplicate checks), or delete a response.

**Can I change tab/header names?**
Keep them as specified unless you also update the Apps Script.

**Can I host the frontend elsewhere?**
Yes—any static host works. Point `WEB_APP_URL` to your Apps Script `/exec`.

---

## Repo structure

```
index.html          # Frontend (GitHub Pages)
backend/Code.gs     # Apps Script backend (paste into the Sheet’s Apps Script)
README.md           # This file
```

---

## License

MIT (or update to your preferred license).
