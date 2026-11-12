# KovsieX Experience

Staff portal and public listener site for **KovsieX**, the student media hub. The pages are hosted on **GitHub Pages**, and the data (staff, bookings, hours, requests, HR) lives in a **Google Sheet** behind a **Google Apps Script** web app.

| Page | File | Who it's for |
|---|---|---|
| Staff portal | `docs/index.html` | Staff, Supervisors and Managers. Sign-in required |
| Listener site | `docs/listen.html` | The public: now playing, send a request, recently played |
| Server | `apps-script/Code.gs` | Runs in Google Apps Script, attached to the KovsieX Google Sheet |

```
GitHub Pages (docs/)  ──POST {fn, args}──▶  Apps Script web app (/exec)  ──▶  Google Sheet
                                                         └──▶  Iono.fm now-playing, Google Drive (show avatars)
```

## 1. Update the Apps Script (one time)

1. Open the KovsieX Google Sheet, then **Extensions → Apps Script**.
2. Replace `Code.gs` with `apps-script/Code.gs` from this repo, and save. This version adds `doPost`, which the GitHub pages talk to.
3. Go to **Deploy → Manage deployments**, click the **pencil**, set **Version: New version**, and click **Deploy**.
   - **Execute as:** Me
   - **Who has access:** Anyone. This only lets the pages reach the sign-in; all data still needs a valid login.
4. Copy the **Web app URL** (it ends in `/exec`).

`apps-script/Index.html` and `apps-script/Public.html` are the Google-hosted versions of the same pages. Keep them in Apps Script as a backup; the old `/exec` link keeps working.

## 2. Point the site at the server

Edit `docs/config.js` and set `apiUrl` to the Web app URL from step 1. It's already set to the current KovsieX deployment. Only change it if you create a **new** deployment, which gets a new URL.

## 3. Turn on GitHub Pages

**Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: `main`, folder `/docs` → Save.**

After a minute the site is live at:

- Staff portal: `https://<your-username>.github.io/<repo-name>/`
- Listener site: `https://<your-username>.github.io/<repo-name>/listen.html`

GitHub Pages is free for **public** repositories. Nothing secret is stored in this repo: passwords and data stay in the Google Sheet.

## Updating later

- **Change how a page looks or works:** edit the file in `docs/` and commit. GitHub Pages updates itself.
- **Change server logic:** paste the new `Code.gs` into Apps Script, then **Manage deployments → pencil → New version → Deploy**. Pick **New version** on the existing deployment so the URL stays the same.

## Live song info on the listener site

The listener site shows the song playing now, with album art, a progress bar and time left. When a song ends it moves to the top of **Recently played**, which keeps the last 5.

- **Which song is playing** comes from **Iono.fm**, which receives song titles with the stream. Set the stream ID in the portal under **Settings → Iono.fm stream ID**. Iono.fm allows one check every 15 seconds, so songs appear within about 15–20 seconds.
- **Faster (about 2 seconds):** if the playout system (Myriad, RadioDJ, …) can send a web request when a song starts, point it at the address and key shown in **Settings → Instant song updates**.
- **Album art and song length** come from Apple Music (iTunes Search). No account is needed.

## Preview without the server

Open `docs/index.html` with `apiUrl` left empty in `config.js`. The portal runs on demo data and a pink banner says so. Any email and password signs in.
