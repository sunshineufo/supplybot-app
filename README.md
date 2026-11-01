# Supplybot app (install on a phone)

A small installable app for Supplybot: its own icon and name, an **Install** button on Android and Add to Home Screen on
iPhone. Opened from the home screen, Supplybot fills the whole screen inside the app, with no browser bar on top.
Supplybot itself still runs from Google Apps Script, so every update you deploy shows up in the app automatically, and
people sign in the same way as in the browser.

Your copy is live at **https://sunshineufo.github.io/supplybot-app/**. Supplybot only lets itself be shown inside an app
at that address (see FRAME_OK at the top of Script.html). If the address ever changes, change it there too.

## Updating your GitHub copy (from the earlier version)
1. In the GitHub repository, upload these files over the old ones: `index.html`, `sw.js`, `config.js`, `README.md`
   (Add file → Upload files → drag them in → Commit changes).
2. Click `config.js` → ✏️ Edit, paste Supplybot's address between the quotes (Supplybot → ⚙ → 📲 Install on your phone → Copy),
   Commit changes.
3. Deploy the new Supplybot version in Apps Script (Code.gs, Index, Styles, Script).
4. On the phone: close the Supplybot app completely and open it again (if it still shows the bar, open it twice: the first
   start fetches the new files).

## Setting it up from scratch
1. Put Supplybot's address in `config.js` (between the quotes).
2. Upload every file in this folder to a GitHub repository and turn on **Settings → Pages → Deploy from a branch → main / (root)**.
3. Open the github.io address on the phone. Android (Chrome): **📲 Install the app**. iPhone (Safari): Share → **Add to Home Screen**.

## Good to know
- If Google ever needs you to sign in again, the app shows **Sign in**: tap it, sign in, come back and tap **Try again**.
- Changing the icon: replace the PNG files (192×192, 512×512, 180×180 for iPhone).
