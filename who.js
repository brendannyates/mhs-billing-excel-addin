/* who.js — "Enter your name" helpers (tested in tests/who.test.mjs). Browser global BEWho.
   Microsoft sign-in stays the default. When it isn't available, the person picks their name once a day:
   the list lives in the workbook's Office document settings ("billingUsers": [{ name, email }]), and today's pick
   is kept on this computer per workbook (localStorage "billingWho:<workbook key>"). Email is the identity, because
   My Tickets matches tickets by the email Microsoft Forms records. */
(function () {
  "use strict";
  const USERS_KEY = "billingUsers";
  const txt = (s) => String(s === null || s === undefined ? "" : s).trim();
  const cleanName = (s) => txt(s).replace(/\s+/g, " ").slice(0, 40).trim();
  const cleanEmail = (s) => { const e = txt(s).toLowerCase(); return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) ? e : ""; };
  /* the saved list, from a document setting (an array, or JSON text): valid emails only, one row per email */
  function readUsers(v) {
    let a = v;
    if (typeof a === "string") { try { a = JSON.parse(a); } catch (e) { a = []; } }
    const out = [], seen = new Set();
    (Array.isArray(a) ? a : []).forEach((u) => {
      const email = cleanEmail(u && u.email), name = cleanName(u && u.name);
      if (email && !seen.has(email)) { seen.add(email); out.push({ name: name || email.split("@")[0], email }); }
    });
    return out;
  }
  /* add or rename one person (matched by email) */
  function addUser(list, u) {
    const email = cleanEmail(u && u.email), name = cleanName(u && u.name);
    const L = readUsers(list);
    if (!email || !name) return L;
    const i = L.findIndex((x) => x.email === email);
    if (i >= 0) L[i] = { name, email }; else L.push({ name, email });
    return L;
  }
  /* the people who've picked on this computer come first, then A to Z; two people with the same name show their email */
  function pickOrder(users, here) {
    const H = (here || []).map((e) => txt(e).toLowerCase());
    const rank = (u) => { const i = H.indexOf(u.email); return i < 0 ? Infinity : i; };
    const L = readUsers(users);
    const names = new Map(); L.forEach((u) => { const k = u.name.toLowerCase(); names.set(k, (names.get(k) || 0) + 1); });
    return L.map((u) => Object.assign({}, u, { here: rank(u) !== Infinity, dup: names.get(u.name.toLowerCase()) > 1 }))
      .sort((a, b) => rank(a) - rank(b) || a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
  }
  /* this computer's memory per workbook: { day, email, used: [emails, latest first] } */
  function whoRec(json) {
    try { const o = JSON.parse(json || "{}"); return { day: txt(o.day), email: cleanEmail(o.email), used: Array.isArray(o.used) ? o.used.map(cleanEmail).filter(Boolean) : [] }; }
    catch (e) { return { day: "", email: "", used: [] }; }
  }
  const whoToday = (rec, day) => (rec && rec.day === day ? rec.email : "");
  const whoPick = (rec, email, day) => ({ day, email, used: [email].concat(((rec && rec.used) || []).filter((e) => e !== email)).slice(0, 20) });
  /* the name screen shows when Microsoft sign-in didn't give an email and nobody on the list is picked for today */
  function needsPick(o) {
    if (!o || o.sso || o.kind === "dataverse" || o.kind === "excel-dialog") return false;
    const t = cleanEmail(o.today);
    return !t || !readUsers(o.users).some((u) => u.email === t);
  }
  /* a short, stable key for a workbook from its URL (or the id saved in it) */
  function wbKeyOf(s) {
    let h = 0x811c9dc5; const t = txt(s).toLowerCase().replace(/[?#].*$/, "");
    for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return t ? "w" + h.toString(36) : "";
  }
  const dayIso = (d) => { d = d || new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); };
  globalThis.BEWho = { USERS_KEY, cleanName, cleanEmail, readUsers, addUser, pickOrder, whoRec, whoToday, whoPick, needsPick, wbKeyOf, dayIso };
})();
