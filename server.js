const express = require("express"), bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken"), path = require("path"), D = require("docx");
const crypto = require("crypto");
const { DatabaseSync } = require("node:sqlite"); // built into Node 22+, nothing to compile
const SECRET = process.env.JWT_SECRET;
if (!SECRET || SECRET.length < 16) { console.error("Set JWT_SECRET (16+ random characters)"); process.exit(1); }
const db = new DatabaseSync(process.env.DB_FILE || "data.db");
db.exec(`create table if not exists users(id integer primary key, email text unique, hash text);
create table if not exists store(user_id integer, k text, v text, primary key(user_id, k));
create table if not exists events(id integer primary key, ts integer, m text, p text, st integer, uid integer, ip text, ms integer);
create index if not exists ev_ts on events(ts)`);
try { db.exec("alter table users add column created integer"); } catch (e) {}
const app = express();
app.set("trust proxy", 1);
app.use(express.json({ limit: "2mb" }));

// Activity log: page visits and API calls only. No resume content, no raw IPs (salted hash), no admin calls.
const logEv = db.prepare("insert into events(ts,m,p,st,uid,ip,ms) values(?,?,?,?,?,?,?)");
app.use((q, s, nx) => {
  const t = Date.now();
  s.on("finish", () => {
    const p = q.path;
    if (p.startsWith("/api/admin") || !(p === "/" || p === "/index.html" || p.startsWith("/api/"))) return;
    try { logEv.run(t, q.method, p, s.statusCode, q.uid || null, crypto.createHash("sha256").update(String(q.ip) + SECRET).digest("hex").slice(0, 10), Date.now() - t); } catch (e) {}
  });
  nx();
});
const purge = () => db.prepare("delete from events where ts<?").run(Date.now() - 30 * 864e5);
purge(); setInterval(purge, 6 * 3600 * 1000).unref();
app.use(express.static(path.join(__dirname, "public")));

const hits = new Map(); setInterval(() => hits.clear(), 15 * 60 * 1000).unref();
const limit = n => (q, s, nx) => { const k = q.ip + q.path, c = (hits.get(k) || 0) + 1; hits.set(k, c);
  c > n ? s.status(429).json({ error: "Too many attempts, try again later" }) : nx(); };
function auth(q, s, nx) {
  try { const p = jwt.verify((q.headers.authorization || "").replace("Bearer ", ""), SECRET); if (!p.uid) throw new Error("x"); q.uid = p.uid; nx(); }
  catch (e) { s.status(401).json({ error: "Please sign in" }); }
}
const token = id => jwt.sign({ uid: Number(id) }, SECRET, { expiresIn: "30d" });

app.post("/api/register", limit(10), (q, s) => {
  const { email = "", password = "" } = q.body;
  if (!/^\S+@\S+\.\S+$/.test(email) || password.length < 8) return s.status(400).json({ error: "Enter a valid email and a password of 8+ characters" });
  try {
    const id = db.prepare("insert into users(email,hash,created) values(?,?,?)").run(email.toLowerCase(), bcrypt.hashSync(password, 10), Date.now()).lastInsertRowid;
    s.json({ token: token(id) });
  } catch (e) { s.status(409).json({ error: "Email already registered" }); }
});
app.post("/api/login", limit(10), (q, s) => {
  const u = db.prepare("select * from users where email=?").get(String(q.body.email || "").toLowerCase());
  if (!u || !bcrypt.compareSync(String(q.body.password || ""), u.hash)) return s.status(401).json({ error: "Wrong email or password" });
  s.json({ token: token(u.id) });
});

const KEYS = ["rb", "jt", "td", "nt", "ev"];
app.get("/api/data", auth, (q, s) => {
  const o = {};
  KEYS.forEach(k => { const r = db.prepare("select v from store where user_id=? and k=?").get(q.uid, k); o[k] = r ? JSON.parse(r.v) : []; });
  s.json(o);
});
app.put("/api/data", auth, (q, s) => {
  const st = db.prepare("insert into store(user_id,k,v) values(?,?,?) on conflict(user_id,k) do update set v=excluded.v");
  KEYS.forEach(k => { if (Array.isArray(q.body[k])) st.run(q.uid, k, JSON.stringify(q.body[k])); });
  s.json({ ok: true });
});

// True .docx from the resume model built by the page (no login needed)
app.post("/api/docx", limit(60), async (q, s) => {
  try {
    const model = Array.isArray(q.body.model) ? q.body.model.slice(0, 500) : [], tpl = String(q.body.tpl || "classic");
    const font = tpl === "executive" ? "Georgia" : "Arial", ac = tpl === "modern" ? "0B4F6C" : "000000", sz = tpl === "compact" ? 20 : 22;
    const centre = tpl === "executive" ? D.AlignmentType.CENTER : D.AlignmentType.LEFT;
    const R = (t, o = {}) => new D.TextRun({ text: String(t || "").slice(0, 2000), font, size: sz, ...o });
    const kids = model.map(m => {
      if (m.t === "h1") return new D.Paragraph({ alignment: centre, children: [R(m.a, { bold: true, size: 40, color: ac })] });
      if (m.t === "m") return new D.Paragraph({ alignment: centre, children: [R(m.a, { size: sz - 2 })] });
      if (m.t === "h3") return new D.Paragraph({ spacing: { before: 200, after: 80 },
        border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: ac, space: 1 } },
        children: [R(tpl === "executive" ? String(m.a).toUpperCase() : m.a, { bold: true, color: ac })] });
      if (m.t === "jt") return new D.Paragraph({ spacing: { before: 100 }, tabStops: [{ type: D.TabStopType.RIGHT, position: 9500 }],
        children: [R(m.a, { bold: true }), new D.TextRun({ children: [new D.Tab(), String(m.b || "")], font, size: sz, bold: true })] });
      if (m.t === "li") return new D.Paragraph({ bullet: { level: 0 }, children: [R(m.a)] });
      return new D.Paragraph({ children: [R(m.a)] });
    });
    const doc = new D.Document({ sections: [{ properties: { page: { margin: { top: 720, bottom: 720, left: 864, right: 864 } } }, children: kids }] });
    const buf = await D.Packer.toBuffer(doc);
    s.set({ "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Content-Disposition": 'attachment; filename="resume.docx"' }).send(buf);
  } catch (e) { s.status(500).json({ error: "Could not build the document" }); }
});

// Optional AI bullet rewrite (needs ANTHROPIC_API_KEY and a signed-in user)
app.post("/api/rewrite", auth, limit(30), async (q, s) => {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return s.status(501).json({ error: "AI rewrite is not configured on this server" });
  const b = String(q.body.bullet || "").slice(0, 500), role = String(q.body.role || "").slice(0, 80);
  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", { method: "POST",
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json" },
      body: JSON.stringify({ model: process.env.ANTHROPIC_MODEL || "claude-sonnet-5-5", max_tokens: 200,
        messages: [{ role: "user", content: `Rewrite this resume bullet for a ${role} resume. Start with a strong past-tense action verb, keep it to one line, keep only the facts given, and use [X] placeholders where a number is missing. Never invent numbers, tools or employers. Return only the bullet.\n\nBullet: ${b}` }] }) });
    const j = await r.json(); if (!r.ok) throw new Error("api");
    s.json({ text: ((j.content && j.content[0] && j.content[0].text) || "").trim() });
  } catch (e) { s.status(502).json({ error: "AI service error" }); }
});

// ---- Admin console API (set ADMIN_USER and ADMIN_PASSWORD, 12+ characters, to enable) ----
const AU = process.env.ADMIN_USER, AP = process.env.ADMIN_PASSWORD;
const eq = (a, b) => crypto.timingSafeEqual(crypto.createHash("sha256").update(String(a)).digest(), crypto.createHash("sha256").update(String(b)).digest());
function adm(q, s, nx) {
  try { if (!jwt.verify((q.headers.authorization || "").replace("Bearer ", ""), SECRET).admin) throw new Error("x"); nx(); }
  catch (e) { s.status(401).json({ error: "Admin sign-in required" }); }
}
app.post("/api/admin/login", limit(5), (q, s) => {
  if (!AU || !AP || AP.length < 12) return s.status(404).json({ error: "Admin is not configured on this server" });
  if (eq(q.body.user || "", AU) && eq(q.body.password || "", AP)) return s.json({ token: jwt.sign({ admin: true }, SECRET, { expiresIn: "8h" }) });
  s.status(401).json({ error: "Wrong admin credentials" });
});
app.get("/api/admin/stats", adm, (q, s) => {
  const d = 864e5, now = Date.now(), day = now - d, week = now - 7 * d, V = "p in ('/','/index.html')";
  const c = (sql, ...a) => db.prepare(sql).get(...a).n;
  s.json({
    users: c("select count(*) n from users"), newUsers7d: c("select count(*) n from users where created>?", week),
    visits24h: c(`select count(*) n from events where ts>? and ${V}`, day), uniq24h: c(`select count(distinct ip) n from events where ts>? and ${V}`, day),
    events24h: c("select count(*) n from events where ts>?", day), errors24h: c("select count(*) n from events where ts>? and st>=500", day),
    failedLogins24h: c("select count(*) n from events where ts>? and p='/api/login' and st=401", day),
    docx24h: c("select count(*) n from events where ts>? and p='/api/docx' and st=200", day),
    rewrites24h: c("select count(*) n from events where ts>? and p='/api/rewrite' and st=200", day),
    series: db.prepare(`select strftime('%Y-%m-%d',ts/1000,'unixepoch') d, sum(${V}) v, count(*) n from events where ts>? group by d order by d`).all(now - 14 * d),
    top: db.prepare("select p,count(*) n from events where ts>? group by p order by n desc limit 8").all(week),
    recent: db.prepare("select ts,m,p,st,uid,ip,ms from events order by id desc limit 100").all(),
    usersList: db.prepare("select u.id,u.email,u.created,(select max(ts) from events e where e.uid=u.id) last from users u order by u.id desc limit 200").all(),
    server: { uptimeMin: Math.round(process.uptime() / 60), memMB: Math.round(process.memoryUsage().rss / 1048576), node: process.version,
      ephemeral: !!process.env.RENDER && !String(process.env.DB_FILE || "").startsWith("/var/data") }
  });
});
app.delete("/api/admin/users/:id", adm, (q, s) => {
  const id = Number(q.params.id);
  db.prepare("delete from store where user_id=?").run(id); db.prepare("delete from users where id=?").run(id);
  s.json({ ok: true });
});

app.listen(process.env.PORT || 3000, () => console.log("Role-Ready running on port " + (process.env.PORT || 3000)));
