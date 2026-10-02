// Outlook calendar, month view (https://outlook.office.com/calendar/view/month). Run in the page's JS tool.
// 1) Check the "Filter applied" menu and note hidden types. 2) Edit CAL_P (or set window.__PARAMS first).
// 3) Run (returns at once), poll window.__evdone. Walks back MONTHS_BACK months from the month on screen.
// Cancelled occurrences are NOT shown by Outlook: infer them from emails/chat. Nothing here is org- or date-specific.
var CAL_P = Object.assign({
  START: '',     // REQUIRED: inclusive start date for parseCalendar
  MONTHS_BACK: 4,          // how many "previous month" clicks (current month + this many back must cover START)
  PAGE_WAIT_MS: 3500,
}, window.__PARAMS || {});
if (!CAL_P.START) throw new Error('Set START (YYYY-MM-DD) in CAL_P or window.__PARAMS before running');
window.__EV = new Map(); window.__evdone = false;
var grab = () => document.querySelectorAll('[aria-label]').forEach(e => {
  const a = e.getAttribute('aria-label') || '';
  if (/, \d{1,2}:\d\d [AP]M to \d{1,2}:\d\d [AP]M, \w+day, \w+ \d{1,2}, \d{4}/.test(a)) window.__EV.set(a, 1);
});
(async () => {
  for (let i = 0; i < CAL_P.MONTHS_BACK; i++) {
    grab();
    const prev = [...document.querySelectorAll('button')].find(b => /previous/i.test(b.getAttribute('aria-label') || '') && !/year/i.test(b.getAttribute('aria-label') || ''));
    if (!prev) { window.__err = 'no previous button'; break; }
    prev.click(); await new Promise(r => setTimeout(r, CAL_P.PAGE_WAIT_MS));
  }
  grab(); window.__evdone = true;
})();

// After __evdone: parse labels -> "YYYY-MM-DD | start | end | title | rest"; dump with the <article> trick (<=8500 chars per chunk).
window.parseCalendar = function () {
  const mon = { January: 0, February: 1, March: 2, April: 3, May: 4, June: 5, July: 6, August: 7, September: 8, October: 9, November: 10, December: 11 };
  const rows = [];
  for (const a of window.__EV.keys()) {
    const m = a.match(/^(.*?), (\d{1,2}:\d\d [AP]M) to (\d{1,2}:\d\d [AP]M), \w+day, (\w+) (\d{1,2}), (\d{4})(.*)$/);
    if (!m) continue;
    const iso = `${m[6]}-${String(mon[m[4]] + 1).padStart(2, '0')}-${m[5].padStart(2, '0')}`;
    if (iso >= CAL_P.START) rows.push([iso, m[2], m[3], m[1], m[7].replace(/^, /, '')].join(' | '));
  }
  return rows.sort();
};
