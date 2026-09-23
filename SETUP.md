# The Haunted Trail: setup

A static web game with no backend. People sign in with M365 (MSAL). Game data lives in three
Microsoft Lists on a SharePoint site, and **SharePoint permissions are the security boundary**.
Nothing enforced only in the browser matters.

```
site/            → deploy this folder (index.html, style.css, app.js, score.js). Contains NO answers.
keeper/          → stays on your laptop. questions.json holds the answers. Never commit or deploy it.
score.test.cjs   → node score.test.cjs
```

Preview the demo (fake data, time travel): `cd site && python3 -m http.server 8765`, then open
`http://localhost:8765/?demo&now=2026-10-06T21:00` (add `&admin` to see the keeper panel).

## How it stays hack-proof(ish)

| List | Players | You | Why |
|---|---|---|---|
| `HQ_Vault` | **no access** | full | Question bank plus answers. Players can't read it, even through the Graph API with their own token. |
| `HQ_Public` | **Read** | full | One item: questions for levels that have opened, answers for levels that have closed, and the leaderboard. |
| `HQ_Submissions` | **Contribute**, *read own items only* | full | Players can't copy each other's answers. |

- Questions appear in `HQ_Public` only once a level opens, and answers only after 8 PM IST. Before that they
  exist only in the Vault.
- Scoring uses SharePoint's own `createdBy` and `lastModified`, so nobody can fake their name or timestamp. A
  submission edited after 8 PM is void. The last submission inside the 9 AM–8 PM window counts.
- Points are computed from raw submissions (`score.js`) and written by you. Players can only read them.

## 1. SharePoint site
Create a site you own, for example `https://uipath.sharepoint.com/sites/HauntedTrail2026` (a team site without
a group is fine). **You should be the only Owner**: site owners can read the Vault.

## 2. Entra app registration (5 min, may need IT)
Entra admin center → App registrations → New:
- Name `Haunted Trail`, single tenant.
- Platform **Single-page application**, redirect URI **`https://parasgera.github.io/haunted-trail/`** (trailing
  slash matters). Add `http://localhost:8765/` for local testing.
- API permissions (delegated): `User.Read`, `Sites.ReadWrite.All`. Microsoft doesn't require admin consent
  for these, but if UiPath has user consent disabled, ask IT to **Grant admin consent**. This is the one
  step that might block you.
- Delegated means the app only ever has the signed-in user's own SharePoint rights, which is why the list
  permissions below are the real lock.
- Copy the **Client ID** and **Tenant ID** into `CONFIG` at the top of `site/app.js`. Set `site` too.

**Staying signed in:** MSAL caches in localStorage and renews silently: refresh token first, then the Entra
SSO session. People sign in once per device. Entra caps SPA refresh tokens at 24 h, but renewal rides on the
normal M365 session, so you won't see a prompt unless UiPath Conditional Access forces one.

## 3. Hosting (done)
Hosted on GitHub Pages at **https://parasgera.github.io/haunted-trail/**. Every push to `main` runs the
scoring test and publishes `site/` (`.github/workflows/pages.yml`). The repo is public, which is safe:
it holds no answers or secrets, and `keeper/` is git-ignored.

## 4. Create the lists
Open `https://parasgera.github.io/haunted-trail/?admin`, sign in, and click **Create the 3 lists**.

## 5. Lock down permissions (do not skip)
On the site, for each list open ⚙️ **List settings → Permissions for this list → Stop inheriting permissions**:
- **HQ_Vault**: remove everyone except you.
- **HQ_Public**: remove the defaults and grant the offsite group (or *Everyone except external users*) **Read**.
- **HQ_Submissions**: grant the same group **Contribute**. Not *Edit*: Edit can change list settings. Then
  **List settings → Advanced settings**:
  - Read access: **Read items that were created by the user**
  - Create and Edit access: **Create items and edit items that were created by the user**

Check it: sign in as a colleague (or an InPrivate window with a test account). `?admin` should say
"Not the keeper".

## 6. Load questions
In `?admin`, use **Import questions** and pick `keeper/questions.json`. You can also edit rows directly in the
`HQ_Vault` list: `Level`, `Seq` (order within the level), `Prompt`, `Options` (one per line, empty for a typed
answer), `Answer` (use `|` for alternatives, and for multiple choice use the exact option text), `Explanation`.
Adding a question on the day is fine. It goes live on the next sync.

## 7. Daily keeper routine
The game only moves forward when your browser syncs, since there's no server. Open `https://parasgera.github.io/haunted-trail/?admin`:
- **after 9 AM** to publish today's level, and
- **after 8 PM** to reveal answers and update the leaderboard.

It syncs on open and every 5 minutes while the tab is open. The keeper table warns ⚠️ when an upcoming level
has no questions yet.

## Changing the schedule
Edit `CONFIG.levels` (dates) and `opensAt` / `closesAt` in `site/app.js`, then redeploy. All times are IST.
