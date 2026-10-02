// Paste into the page's JS tool while an Outlook web search result list is showing.
// Edit OUTLOOK_P first (or set window.__PARAMS before pasting). Nothing else in this file is org- or date-specific.
// Step 1: run (returns at once). Step 2: poll window.__done / window.__R.size.
// Step 3: window.dumpOutlook() writes the filtered rows into an <article>; read it with the get-page-text action in
//         chunks of at most OUTLOOK_P.CHUNK_CHARS characters (the JS tool returns only ~800 chars directly).
var OUTLOOK_P = Object.assign({
  START: '',          // REQUIRED: inclusive start date, YYYY-MM-DD
  ANCESTOR_LEVELS: 5,           // scroll container = this many parents above the first row (check in the page if Outlook changes)
  SCROLL_PX: 400, PAUSE_MS: 600, STABLE_STEPS: 8, MAX_STEPS: 900,
  KEY_CHARS: 200,               // dedupe key length of the aria-label
  LABEL_CHARS: 400,             // chars of each label kept in memory
  ROW_CHARS: 300,               // chars of each row dumped
  CHUNK_CHARS: 8500,
}, window.__PARAMS || {});
if (!OUTLOOK_P.START) throw new Error('Set START (YYYY-MM-DD) in OUTLOOK_P or window.__PARAMS before running');
window.__PARAMS_OUTLOOK = OUTLOOK_P;
window.__done = false; window.__R = new Map();
(async () => {
  let sc = document.querySelector('div[data-convid]');
  for (let i = 0; i < OUTLOOK_P.ANCESTOR_LEVELS; i++) sc = sc.parentElement;
  sc.scrollTop = 0; await new Promise(r => setTimeout(r, OUTLOOK_P.PAUSE_MS + 200));
  let stable = 0, last = 0, lastH = 0;
  for (let i = 0; i < OUTLOOK_P.MAX_STEPS && stable < OUTLOOK_P.STABLE_STEPS; i++) {
    document.querySelectorAll('div[data-convid]').forEach(r => {
      const a = (r.getAttribute('aria-label') || '').replace(/\s+/g, ' ');
      const k = a.slice(0, OUTLOOK_P.KEY_CHARS);
      if (!window.__R.has(k)) window.__R.set(k, a.slice(0, OUTLOOK_P.LABEL_CHARS));
    });
    sc.scrollTop += OUTLOOK_P.SCROLL_PX; await new Promise(r => setTimeout(r, OUTLOOK_P.PAUSE_MS));
    if (window.__R.size === last && sc.scrollHeight === lastH) stable++; else { stable = 0; last = window.__R.size; lastH = sc.scrollHeight; }
    const m = [...window.__R.values()].slice(-1)[0].match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);   // stop once older than the start
    if (m && i > 5 && `${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}` < OUTLOOK_P.START) break;
  }
  window.__done = true;
})();

// Resolve a row label to YYYY-MM-DD. Weekday-only labels (last 7 days) need today's date.
window.dumpOutlook = function (TODAY = new Date()) {
  const wd = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };
  const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const rows = [];
  for (let a of window.__R.values()) {
    a = a.replace(/^(Unread |Collapsed |Has attachments )+/, '');
    let m = a.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/), d = null;
    if (m) d = `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
    else if ((m = a.match(/\b(Sun|Mon|Tue|Wed|Thu|Fri|Sat)(?: (\d{1,2})\/(\d{1,2}))? \d{1,2}:\d\d [AP]M/))) {
      if (m[2]) d = `${TODAY.getFullYear()}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
      else { const t = new Date(TODAY); while (t.getDay() !== wd[m[1]]) t.setDate(t.getDate() - 1); d = iso(t); }
    }
    if (d && d >= OUTLOOK_P.START) rows.push(d + ' | ' + a.slice(0, OUTLOOK_P.ROW_CHARS));
  }
  rows.sort();
  document.getElementById('__dump')?.remove();
  const el = document.createElement('article'); el.id = '__dump';
  el.style.cssText = 'position:fixed;inset:0;z-index:99999;background:#fff;color:#000;overflow:auto;white-space:pre-wrap;font-size:9px';
  document.body.appendChild(el); el.textContent = rows.join('\n');
  return rows.length + ' rows, ' + el.textContent.length + ' chars (split into <=' + OUTLOOK_P.CHUNK_CHARS + '-char chunks if larger)';
};
