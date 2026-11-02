# Teams Next Meeting Tile

A native Windows (.NET 9 WPF) dark-mode overlay tile designed to dock directly inside the **Microsoft Teams** desktop sidebar (right above **Quick views**) and display:
- **Meeting Title**
- **Start Time** (e.g., `Start: 1:30 PM` or `Start: Mon 9:00 AM`)
- **Dynamic Countdown Badge**:
  - Over 2 days (`> 2d 0h 0m`): `#d` (e.g., `3d`)
  - Between 1 and 2 days: `#d #h #m` (e.g., `1d 4h 20m`)
  - Under 1 day (`>= 1h`): `#h #m` (e.g., `1h 43m`)
  - Under 1 hour: `#m` (e.g., `43m`, or `0m (45s)` under 1 minute)
- **Full Timezone & Daylight Saving Support**: Converts recurring (`RRULE`) and single events across timezones (e.g. AEDT ↔ AEST) per occurrence date
- **One-Click Teams Join button** when a Teams meeting link is present

---

## How to Launch
Double-click:
- [`Launch-TeamsNextMeetingTile.bat`](./Launch-TeamsNextMeetingTile.bat)
- Or run the executable directly: [`bin\Release\net9.0-windows\TeamsNextMeetingTile.exe`](./bin/Release/net9.0-windows/TeamsNextMeetingTile.exe)

---

## Zero-Export Calendar Sync (3 Built-in Engines)
1. **Microsoft 365 / New Outlook Live Sync (WebView2 SSO) — *Default***
   - Connects in the background to `https://outlook.office.com/calendar/view/workweek` using Edge WebView2 with Windows Single Sign-On (`AllowSingleSignOnUsingOSPrimaryAccount`).
   - Intercepts Outlook Web's live calendar data (`GetCalendarView` / Graph JSON + rendered calendar DOM events) automatically in the background.
   - If your organization prompts you to select your work account the very first time, click **`Sign In`** on the tile (or click **`Open M365 Calendar Window`** in `⚙ Settings`), click your account once, and click **`Run in Background (Hide)`**.
2. **Optional One-Time Live ICS Feed URL**
   - In New Outlook (`Settings ⚙ → Calendar → Shared calendars → Publish a calendar`), you can generate a permanent live `.ics` URL and paste it once in `⚙ Settings`. The app polls the URL every 60 seconds and expands recurring `RRULE` meetings automatically.
3. **Classic Outlook COM (MAPI)**
   - If Classic Outlook (`OUTLOOK.EXE`) is running, the app also reads upcoming calendar items directly via local MAPI.

---

## Positioning & Controls
- **Default Dock Position:** Automatically tracks the Microsoft Teams window and docks at `(OffsetX = 52, OffsetY = 122)`, right above **Quick views** in the Chat sidebar.
- **Drag to Adjust:** Drag the tile anywhere on the Teams window—it remembers its exact offset relative to Teams and moves with the Teams window across monitors.
- **Smart Auto-Hide:** Stays visible when Teams is visible (even on a secondary monitor while you work on your primary monitor), and hides automatically if Teams is minimized or covered by another window over that area.
- **Collapse (`▾`):** Collapses into a slim single-line bar (`1:30 PM • 0d 1h 25m • Title`) if you want to uncover items directly beneath it.
- **System Tray Icon:** Right-click or double-click the purple clock icon in the Windows System Tray at any time to open Settings, trigger a manual sync, or snap the tile back to its default spot.
