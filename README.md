# Supplybot app (install on a phone)

This folder is a small installable app ("PWA") for Supplybot: its own icon, name and splash, an **Install** button on
Android, and Add to Home Screen on iPhone. When it's opened from the home screen it goes straight into Supplybot.
Supplybot itself still runs from Google Apps Script, so every update you deploy shows up in the app automatically,
and people sign in the same way as in the browser.

## 1. Put in Supplybot's address
1. In Supplybot tap **⚙ → 📲 Install on your phone → Copy** (it copies Supplybot's address).
2. Open `index.html` in any text editor and paste it into this line, between the quotes:
   `const APP_URL = 'PASTE_SUPPLYBOT_ADDRESS_HERE';`

## 2. Put the folder online (free)
An installable app needs its own https address. The easiest free option is **GitHub Pages**:
1. Create a GitHub repository (e.g. `supplybot-app`) and upload every file in this folder.
2. In the repository: **Settings → Pages → Build and deployment → Deploy from a branch → main / (root) → Save**.
3. After a minute it's live at `https://<your-github-name>.github.io/supplybot-app/`.

(Netlify, Cloudflare Pages or Firebase Hosting work the same way: upload the folder.)

## 3. Install it on a phone
Open that address on the phone:
- **Android (Chrome):** tap **📲 Install the app** (or ⋮ → Install app).
- **iPhone / iPad (Safari):** Share → **Add to Home Screen** → Add.

Share the address with the team so they can install it too.

## Good to know
- The app opens the real Supplybot, so it's always the latest version and uses the browser's Google sign-in.
- On iPhone, Apple may open Supplybot in a Safari view inside the app. If you'd rather skip this package on iPhone, you can
  also Add to Home Screen straight from Supplybot itself (⚙ → 📲 Install on your phone shows how).
- Changing the icon: replace the PNG files (192×192, 512×512, 180×180 for iPhone) and bump `CACHE` in `sw.js`.
