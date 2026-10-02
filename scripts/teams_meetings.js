// Teams web: pull "Meeting started / Meeting ended" cards from the open chat (real meeting durations).
// Usage: open a meeting chat, run this, then `await window.__scanMeet()` (background it: it scrolls ~1 s per step).
// To sweep all chats, loop over leaf [role="treeitem"] rows exactly like teams_collect.js and attribute events by
// document.title AFTER the scan (the list reorders while you click, and the title lags by a few seconds).
var MEET_P = Object.assign({
  START: '',        // REQUIRED: inclusive start date, YYYY-MM-DD; scrolling stops once separators older than this appear
  SCROLL_PAUSE_MS: 900, MAX_SCROLL_STEPS: 120, STABLE_STEPS: 3,
  MAX_CARD_CHARS: 250,
  ROUND_SECONDS_UP_AT: 30,    // seconds at or above this round the minute up
}, window.__PARAMS || {});
if (!MEET_P.START) throw new Error('Set START (YYYY-MM-DD) in MEET_P or window.__PARAMS before running');
var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
// words that appear in date separators older than START (earlier months of that year, and the previous year)
var _s = new Date(MEET_P.START + 'T00:00:00');
var OLDER = new RegExp('\\b(' + [...MONTHS.slice(0, _s.getMonth()), String(_s.getFullYear() - 1)].join('|') + ')\\b');

window.__scanMeet = async function () {
  const vp = document.querySelector('[data-tid="message-pane-list-viewport"]'); if (!vp) return [];
  let stable = 0, last = -1;
  for (let i = 0; i < MEET_P.MAX_SCROLL_STEPS; i++) {
    vp.scrollTop = 0; await new Promise(r => setTimeout(r, MEET_P.SCROLL_PAUSE_MS));
    const n = vp.querySelectorAll('[data-tid="chat-pane-item"], [role="separator"]').length;
    if (n === last) { if (++stable >= MEET_P.STABLE_STEPS) break; } else { stable = 0; last = n; }
    const seps = [...vp.querySelectorAll('[role="separator"]')].map(s => s.innerText.trim());
    if (seps.some(s => OLDER.test(s))) break;
  }
  const out = []; const seen = new Set();
  vp.querySelectorAll('[data-tid="chat-pane-item"]').forEach(e => {
    const t = (e.innerText || '').replace(/\s+/g, ' ').trim();
    if (/^Meeting (started|ended)/.test(t) && t.length < MEET_P.MAX_CARD_CHARS && !seen.has(t)) { seen.add(t); out.push(t); }
  });
  return out;
};

// Parse one card: "Meeting ended: at 8/5 4:42 PM after 42 minutes 17 seconds ..." / "Meeting started at Wednesday 12:59 PM"
// Dates are M/D, M/D/YYYY (older than a year ago) or a bare weekday (last 7 days: resolve against today).
window.parseMeetCard = function (t, today = new Date()) {
  const wd = { Sunday: 0, Monday: 1, Tuesday: 2, Wednesday: 3, Thursday: 4, Friday: 5, Saturday: 6 };
  const m = t.match(/^Meeting (started|ended)(?::)? at (?:(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?|(\w+day)) (\d{1,2}:\d\d [AP]M)(?: after (?:(\d+) hours? )?(?:(\d+) minutes? )?(?:(\d+) seconds?)?)?/);
  if (!m) return null;
  let d;
  if (m[2]) d = `${m[4] || today.getFullYear()}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  else { const x = new Date(today); while (x.getDay() !== wd[m[5]]) x.setDate(x.getDate() - 1); d = `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`; }
  const mins = m[1] === 'ended' ? (+m[7] || 0) * 60 + (+m[8] || 0) + ((+m[9] || 0) >= MEET_P.ROUND_SECONDS_UP_AT ? 1 : 0) : null;
  return { date: d, time: m[6], kind: m[1], minutes: mins };
};
