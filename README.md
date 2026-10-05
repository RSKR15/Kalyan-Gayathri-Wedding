# Saikalyan & Gayathri — Wedding Invite

Digital wedding invitation with RSVP.

- **Mehendi** — Wednesday, November 18, 2026 · Evening · Serene Springs Farm, 121 S Branch Rd, Hillsborough Township, NJ 08844
- **Wedding (Muhurtham)** — Friday, November 20, 2026 · 9:00 AM · Sri Venkateswara Temple, 1 Balaji Temple Dr, Bridgewater, NJ 08807
- RSVP by Sunday, November 1, 2026

## Publish with GitHub Pages
Settings → Pages → Source: **Deploy from a branch** → Branch: `main`, folder `/ (root)` → Save.
The invite will be live at `https://rskr15.github.io/Kalyan-Gayathri-Wedding/` after a minute or two.

## Connect RSVPs to a Google Sheet
1. Create a Google Sheet → **Extensions → Apps Script**.
2. Replace the code with the contents of `apps-script.gs` and save.
3. **Deploy → New deployment → Web app** — Execute as: *Me*, Who has access: *Anyone* → Deploy and authorize.
4. Copy the Web app URL (ends in `/exec`) and paste it into `index.html` in place of `PASTE_YOUR_APPS_SCRIPT_URL_HERE`.

Responses appear in the **RSVPs** tab, one row per phone number (a guest who replies again updates their row).
