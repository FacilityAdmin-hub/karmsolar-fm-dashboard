# Security review: KarmSolar FM dashboard

Reviewed 2026-11-15. Covers `index.html`, the ticket pages, `worker.js`, `sync-script/` and the GitHub Actions workflows.

## Needs action by an admin (cannot be fixed in code)

### 1. Company data is public. Critical.
The repository is **public**, and `index.html` contains the dashboard data in plain text: payments, vendors, budget, contracts, rents, staff attendance and assets. The public GitHub Issues contain requester names and emails. Anyone can read this on GitHub, or by viewing the source of the GitHub Pages / Vercel site. The Microsoft sign-in in the dashboard does **not** protect it, because the data is already in the page before anyone signs in.

Recommended:
1. Make the repository **private** (GitHub → Settings → General → Danger zone → Change visibility).
   - GitHub Pages on a private repo needs a paid plan. Host on Vercel instead, with **Vercel Authentication / Password Protection** turned on, or put the site behind **Cloudflare Access** limited to `@karmsolar.com`.
   - Private repos have a monthly Actions-minutes quota. The two sync workflows run every 5 minutes, so check usage after the switch.
2. Older data is still in the git history (every auto-sync commit). If the data is sensitive, contact GitHub Support to purge cached views after making the repo private.

### 2. The ticket proxy accepts requests from anyone. High.
The deployed Cloudflare Worker (`karmsolar-ticket-proxy`) has no authentication. Anyone who finds the URL can create, edit, delete or re-status tickets, and add comments or attachments, with a simple script.

Recommended, in the Cloudflare dashboard:
1. Port the hardening from `worker.js` in this repo into the deployed worker:
   - origin allow-list
   - 32 KB body limit
   - field allow-list with length limits
   - company-domain-only requester emails
   - no GitHub error text echoed back to the caller
2. Better: protect the board and its worker routes with **Cloudflare Access** (free for up to 50 users), so only signed-in KarmSolar staff can call the mutating actions.
3. Add a **rate-limiting rule** for the worker route (Security → WAF → Rate limiting), e.g. 30 requests per minute per IP.
4. Make sure the worker's `GITHUB_TOKEN` is a **fine-grained token** scoped to this repo only, with **Issues: read & write** and nothing else. Rotate it if it was ever shared.

## Fixed in this change

| Risk | Where | Fix |
|---|---|---|
| Anyone with a GitHub account could open a `Ticket:` / `Delete:` / `UpdateStatus:` issue and write to or delete from the FM Tickets workbook, bypassing the worker entirely | `.github/workflows/ticket-*.yml` | Jobs now run only when the issue was opened by the repository owner (the account the worker posts as) |
| Spreadsheet formula injection: a ticket title like `=HYPERLINK(...)` was stored as a live Excel formula | `sync-script/write-ticket.js` | Values starting with `= + - @` are stored as text; all fields are length-limited |
| The form could be abused to send company-branded email from the FM mailbox to any outside address | `write-ticket.js`, `update-status.js` | Requester emails go only to `@karmsolar.com` (override with `ALLOWED_EMAIL_DOMAINS`) |
| HTML injection in requester emails (name and title were not escaped) | `write-ticket.js`, `update-status.js` | All values escaped |
| Unvalidated ticket IDs | `delete-ticket.js`, `update-status.js` | Strict ID format check |
| Cross-site scripting (XSS): the ticket `Priority` was inserted into the page unescaped | `ticket-board.html` | Escaped |
| `javascript:` links possible in attachment URLs | `ticket-board.html` | Only `http(s)` links are rendered |
| Incomplete HTML escaping helper (missing `'` and `>`) | `index.html` `esc()` | Escapes all five characters |
| Outdated SheetJS (`xlsx` 0.18.5 has known CVEs: prototype pollution, ReDoS) | workflows, `index.html` | Upgraded to the official SheetJS 0.20.3 build |
| No browser security headers | `vercel.json` (new) | `nosniff`, `X-Frame-Options`, `Referrer-Policy`, `Permissions-Policy`, HSTS |
| No origin, size or field checks in the reference worker | `worker.js` | Hardened (see item 2 above to apply it to the deployed worker) |

## Checked and OK
- No passwords, tokens or client secrets are committed. Azure and GitHub credentials live in GitHub Actions secrets.
- Workflows pass issue content through environment variables, not `${{ }}` inside `run:` scripts, so there is no shell injection.
- Ticket status values are checked against a fixed list.
