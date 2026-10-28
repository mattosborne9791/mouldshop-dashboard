/*
 * Sales inbox response-time logic for sales@mouldshop.com.
 * Pure functions only (no DOM, no network) so they can be tested in Node.
 *
 * Business hours (UK time): Mon–Thu 08:30–17:00, Fri 08:30–16:00.
 * Weekends and England & Wales bank holidays are closed.
 */
(function (root) {
  "use strict";

  const CONFIG = {
    ourDomain: "mouldshop.com",
    mailbox: "sales@mouldshop.com",
    slaMinutes: 120, // 2 business hours
    timeZone: "Europe/London",
    // [open, close] in minutes after midnight, by JS weekday (0 = Sunday)
    hours: {
      1: [510, 1020], // Mon 08:30–17:00
      2: [510, 1020], // Tue
      3: [510, 1020], // Wed
      4: [510, 1020], // Thu
      5: [510, 960],  // Fri 08:30–16:00
    },
    // England & Wales bank holidays
    holidays: [
      "2026-01-01", "2026-04-03", "2026-04-06", "2026-05-04", "2026-05-25",
      "2026-08-31", "2026-12-25", "2026-12-28",
      "2027-01-01", "2027-03-26", "2027-03-29", "2027-05-03", "2027-05-31",
      "2027-08-30", "2027-12-27", "2027-12-28",
    ],
    // Senders that are never customers waiting on a reply.
    excludeDomains: ["socialintents.com", "mailchimp.com", "mcdlv.net", "hubspotemail.net", "linkedin.com"],
    automatedLocalPart: /(^|[._+-])(no-?reply|do-?not-?reply|donotreply|notifications?|notify|mailer-daemon|postmaster|bounces?|newsletters?|news|marketing|mailer|alerts?|automated|updates?|hello)([._+-]|$)/i,
    automatedSubject: /^\s*(automatic reply|auto(matic)?[- ]?reply|autoreply|out of (the )?office|undeliverable|delivery (status notification|has failed|failure)|read:|accepted:|declined:|tentative:|mail delivery)/i,
    // An Outlook category on a customer email that means "doesn't need a reply".
    noReplyCategory: /no\s*reply/i,
  };

  // ---------- time zone helpers ----------
  const dtf = new Intl.DateTimeFormat("en-GB", {
    timeZone: CONFIG.timeZone, hourCycle: "h23",
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", weekday: "short",
  });
  const WD = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };

  function londonParts(ms) {
    const p = {};
    for (const part of dtf.formatToParts(new Date(ms))) p[part.type] = part.value;
    return {
      y: +p.year, m: +p.month, d: +p.day, h: +p.hour, min: +p.minute, s: +p.second,
      weekday: WD[p.weekday],
      dateKey: `${p.year}-${p.month}-${p.day}`,
    };
  }
  function offsetMs(ms) {
    const p = londonParts(ms);
    return Date.UTC(p.y, p.m - 1, p.d, p.h, p.min, p.s) - Math.floor(ms / 1000) * 1000;
  }
  // UK local wall-clock time -> UTC milliseconds
  function londonToUtc(y, m, d, minutesOfDay) {
    const guess = Date.UTC(y, m - 1, d, 0, minutesOfDay);
    let t = guess - offsetMs(guess);
    t = guess - offsetMs(t);
    return t;
  }

  const holidaySet = new Set(CONFIG.holidays);

  // Opening window for the UK calendar day containing ms: [openUtc, closeUtc] or null
  function dayWindow(y, m, d) {
    const weekday = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
    const key = `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    const hrs = CONFIG.hours[weekday];
    if (!hrs || holidaySet.has(key)) return null;
    return [londonToUtc(y, m, d, hrs[0]), londonToUtc(y, m, d, hrs[1])];
  }
  function nextDay(y, m, d) {
    const t = new Date(Date.UTC(y, m - 1, d + 1));
    return [t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()];
  }

  // Business minutes between two instants
  function businessMinutes(startMs, endMs) {
    if (!(endMs > startMs)) return 0;
    let { y, m, d } = londonParts(startMs);
    let total = 0;
    for (let i = 0; i < 400; i++) {
      const w = dayWindow(y, m, d);
      if (w) {
        if (w[0] >= endMs) break;
        const a = Math.max(w[0], startMs), b = Math.min(w[1], endMs);
        if (b > a) total += (b - a) / 60000;
      } else if (londonToUtc(y, m, d, 0) >= endMs) break;
      [y, m, d] = nextDay(y, m, d);
    }
    return total;
  }

  // The instant at which `minutes` business minutes have passed since startMs
  function addBusinessMinutes(startMs, minutes) {
    let { y, m, d } = londonParts(startMs);
    let left = minutes;
    for (let i = 0; i < 400; i++) {
      const w = dayWindow(y, m, d);
      if (w) {
        const a = Math.max(w[0], startMs);
        if (w[1] > a) {
          const avail = (w[1] - a) / 60000;
          if (avail >= left) return a + left * 60000;
          left -= avail;
        }
      }
      [y, m, d] = nextDay(y, m, d);
    }
    return NaN;
  }

  // ---------- message classification ----------
  const lower = (s) => (s || "").toLowerCase().trim();
  const addr = (r) => lower(r && r.emailAddress && r.emailAddress.address);
  const isOurs = (a) => a.endsWith("@" + CONFIG.ourDomain);

  function isAutomatedSender(a) {
    if (!a) return true;
    const [local, domain = ""] = a.split("@");
    if (CONFIG.excludeDomains.some((x) => domain === x || domain.endsWith("." + x))) return true;
    return CONFIG.automatedLocalPart.test(local);
  }

  // Normalise a Graph message into what the calculations need
  function normalise(msg) {
    const from = addr(msg.from);
    const recipients = [...(msg.toRecipients || []), ...(msg.ccRecipients || [])].map(addr).filter(Boolean);
    const internal = isOurs(from);
    return {
      id: msg.id,
      conversationId: msg.conversationId || msg.id,
      subject: msg.subject || "(no subject)",
      from,
      fromName: (msg.from && msg.from.emailAddress && msg.from.emailAddress.name) || from,
      recipients,
      internal,
      at: Date.parse(internal ? (msg.sentDateTime || msg.receivedDateTime) : (msg.receivedDateTime || msg.sentDateTime)),
      categories: msg.categories || [],
      preview: msg.bodyPreview || "",
      webLink: msg.webLink || "",
      skip: !!msg.isDraft || !!msg.isJunk,
    };
  }

  function isCustomerEmail(n) {
    if (n.skip || n.internal) return false;
    if (isAutomatedSender(n.from)) return false;
    if (CONFIG.automatedSubject.test(n.subject)) return false;
    if (n.categories.some((c) => CONFIG.noReplyCategory.test(c))) return false;
    return true;
  }

  function isReply(n) {
    // Sent by someone at Mouldshop to at least one outside address
    return !n.skip && n.internal && n.recipients.some((r) => !isOurs(r));
  }

  // Who answered: a named colleague, or the sign-off name in a reply sent as sales@
  function responderOf(n) {
    if (n.from !== CONFIG.mailbox) return n.fromName || n.from;
    const lines = n.preview.replace(/[​-‍﻿]/g, "").split(/\r?\n/).map((l) => l.trim());
    const sig = lines.findIndex((l) => /^sales team$/i.test(l));
    const searchTo = sig === -1 ? lines.length : sig;
    for (let i = searchTo - 1; i >= 0; i--) {
      const l = lines[i];
      if (!l) continue;
      if (/^[A-Z][a-zA-Z'-]{1,20}( [A-Z][a-zA-Z'-]{1,20})?$/.test(l) &&
          !/^(thanks|thank you|regards|kind regards|best regards|many thanks|cheers|hi|hello|dear)$/i.test(l)) return l;
      if (sig !== -1) break; // only trust the line right above the signature
    }
    return "Sales Team (unsigned)";
  }

  /*
   * Build one row per customer email that needed a reply.
   * Consecutive customer emails in a thread before we answer count once,
   * from the first of them. The reply is the next email from Mouldshop in
   * that conversation sent to an outside address.
   */
  function buildTurns(rawMessages, nowMs) {
    const byConv = new Map();
    for (const raw of rawMessages) {
      const n = normalise(raw);
      if (!Number.isFinite(n.at)) continue;
      if (!byConv.has(n.conversationId)) byConv.set(n.conversationId, []);
      byConv.get(n.conversationId).push(n);
    }
    const turns = [];
    for (const msgs of byConv.values()) {
      msgs.sort((a, b) => a.at - b.at);
      let pending = null;
      for (const n of msgs) {
        if (isCustomerEmail(n)) {
          if (!pending) pending = n;
        } else if (pending && isReply(n) && n.at >= pending.at) {
          turns.push(makeTurn(pending, n, nowMs));
          pending = null;
        }
      }
      if (pending) turns.push(makeTurn(pending, null, nowMs));
    }
    return turns.sort((a, b) => a.receivedAt - b.receivedAt);
  }

  function makeTurn(inbound, reply, nowMs) {
    const end = reply ? reply.at : nowMs;
    const minutes = businessMinutes(inbound.at, end);
    const p = londonParts(inbound.at);
    return {
      id: inbound.id,
      subject: inbound.subject,
      customer: inbound.fromName,
      customerEmail: inbound.from,
      webLink: inbound.webLink,
      receivedAt: inbound.at,
      receivedWeekday: p.weekday,
      receivedHour: p.h,
      answered: !!reply,
      repliedAt: reply ? reply.at : null,
      responder: reply ? responderOf(reply) : null,
      minutes, // business minutes to reply, or waiting so far
      dueAt: addBusinessMinutes(inbound.at, CONFIG.slaMinutes),
      withinSla: minutes <= CONFIG.slaMinutes,
    };
  }

  // ---------- summaries ----------
  function median(xs) {
    if (!xs.length) return NaN;
    const s = [...xs].sort((a, b) => a - b), mid = s.length >> 1;
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  }
  const mean = (xs) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : NaN);

  function summarise(turns) {
    const answered = turns.filter((t) => t.answered);
    const mins = answered.map((t) => t.minutes);
    return {
      received: turns.length,
      answered: answered.length,
      avg: mean(mins),
      median: median(mins),
      slaRate: answered.length ? answered.filter((t) => t.withinSla).length / answered.length : NaN,
      waiting: turns.filter((t) => !t.answered).length,
      overdue: turns.filter((t) => !t.answered && !t.withinSla).length,
    };
  }

  // Monday 00:00 UK of the week containing ms, as "YYYY-MM-DD"
  function weekKey(ms) {
    const p = londonParts(ms);
    const back = (p.weekday + 6) % 7;
    const t = new Date(Date.UTC(p.y, p.m - 1, p.d - back));
    return t.toISOString().slice(0, 10);
  }

  function byWeek(turns) {
    const map = new Map();
    for (const t of turns) {
      const k = weekKey(t.receivedAt);
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(t);
    }
    return [...map.entries()].sort(([a], [b]) => (a < b ? -1 : 1))
      .map(([week, ts]) => ({ week, ...summarise(ts) }));
  }

  function byResponder(turns) {
    const map = new Map();
    for (const t of turns) {
      if (!t.answered) continue;
      if (!map.has(t.responder)) map.set(t.responder, []);
      map.get(t.responder).push(t);
    }
    return [...map.entries()].map(([name, ts]) => ({ name, ...summarise(ts) }))
      .sort((a, b) => b.answered - a.answered);
  }

  // Weekday (1–5) × hour grid of customer emails and how quickly they were answered
  function byDayHour(turns) {
    const cells = {};
    for (const t of turns) {
      const key = t.receivedWeekday + ":" + t.receivedHour;
      (cells[key] = cells[key] || []).push(t);
    }
    const out = {};
    for (const [k, ts] of Object.entries(cells)) out[k] = summarise(ts);
    return out;
  }

  const api = {
    CONFIG, londonParts, londonToUtc, businessMinutes, addBusinessMinutes,
    normalise, isCustomerEmail, isReply, responderOf, buildTurns,
    summarise, byWeek, byResponder, byDayHour, median, mean, weekKey,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.SalesResponse = api;
})(typeof window !== "undefined" ? window : globalThis);
