# Role-Ready (Node.js back end)

Resume builder, job tracker, tasks, notes, calendar and insights, with accounts, cloud sync, real .docx export and optional AI bullet rewrite.

## Run
```
npm install
export JWT_SECRET="a-long-random-string"      # required
export ANTHROPIC_API_KEY="sk-ant-..."         # optional, enables AI rewrite
npm start                                     # http://localhost:3000
```
Windows PowerShell: `$env:JWT_SECRET="..."`

## Notes
- Data: SQLite file `data.db` (set DB_FILE to change). Back it up regularly.
- Passwords are hashed with bcrypt. Serve behind HTTPS in production (Render, Railway, Fly.io, or a VPS with Nginx).
- Sync is manual: "Upload to cloud" replaces the cloud copy, "Download from cloud" replaces the browser copy.
- The page works without the server too (localStorage only); .docx, accounts and AI need the server.

## Admin console
Set `ADMIN_USER` and `ADMIN_PASSWORD` (12+ characters), then open `/admin.html`.
It shows visits, unique visitors, API calls, new users, Word exports, AI rewrites, failed logins, server errors, a 14-day chart, top requests, a user list (with delete) and the latest 100 events.
Logged: path, method, status, time, user id and a salted hash of the IP. Never resume content or raw IPs. Events older than 30 days are deleted. Mention this in your privacy policy.

## Free hosting warning (Render free)
The free plan has a temporary disk and sleeps after 15 minutes idle, so `data.db` (accounts and activity log) is erased on restart or redeploy. Fine for a demo; for real users use a persistent database or a paid disk.
