// Paste into the page's JS tool on teams.cloud.microsoft after the app has finished loading (~20 s).
// Click "See more" in the chat list first. Edit TEAMS_P (or set window.__PARAMS first). Then run, poll window.__chdone,
// inspect window.__CH. Nothing here is org- or date-specific.
var TEAMS_P = Object.assign({
  START: '',              // REQUIRED: inclusive start date, YYYY-MM-DD, interpreted as local midnight in the browser's zone
  SCROLL_PAUSE_MS: 900, MAX_SCROLL_STEPS: 300, STABLE_STEPS: 3,
  CLICK_WAIT_MS: 3500,              // wait after clicking a chat row before reading it
  MAX_TEXT: 700,                    // chars kept per message
  ME_MARK: 'ChatMyMessage',         // class marker on my own messages
  ME_LABEL: 'ME',                   // author label written for my messages (map to your name in the config)
  SKIP_ROWS: '^(Copilot|Quick views|Mentions|Drafts|Favorites|Chats|Teams and channels|See all your teams|Communities|Join communities|See more|See less)',
}, window.__PARAMS || {});
if (!TEAMS_P.START) throw new Error('Set START (YYYY-MM-DD) in TEAMS_P or window.__PARAMS before running');
window.__STARTMS = new Date(TEAMS_P.START + 'T00:00:00').getTime();

window.__grab = async function () {
  const vp = document.querySelector('[data-tid="message-pane-list-viewport"]'); if (!vp) return null;
  let stable = 0, lastCount = -1;
  for (let i = 0; i < TEAMS_P.MAX_SCROLL_STEPS; i++) {
    vp.scrollTop = 0; await new Promise(r => setTimeout(r, TEAMS_P.SCROLL_PAUSE_MS));
    const mids = [...vp.querySelectorAll('[data-mid]')].map(e => +e.getAttribute('data-mid')).filter(x => x > 1e12);
    if (mids.length && Math.min(...mids) < window.__STARTMS) break;       // reached before the start date
    if (mids.length === lastCount) { if (++stable >= TEAMS_P.STABLE_STEPS) break; } else { stable = 0; lastCount = mids.length; }
  }
  const out = []; let author = '';
  vp.querySelectorAll('[data-tid="chat-pane-item"]').forEach(it => {
    const m = it.querySelector('[data-tid="chat-pane-message"]'); const mid = +(m?.getAttribute('data-mid') || 0);
    const me = it.innerHTML.slice(0, 2500).includes(TEAMS_P.ME_MARK);
    const a = it.querySelector('[data-tid="message-author-name"]')?.innerText.trim();
    if (a) author = a;                                                      // later messages by the same author omit the name
    if (!(mid > 1e12)) return;
    out.push({ mid, author: me ? TEAMS_P.ME_LABEL : (a || author), text: (m.innerText || '').replace(/\s+/g, ' ').trim().slice(0, TEAMS_P.MAX_TEXT) });
  });
  return out;
};

window.__CH = []; window.__chdone = false; window.__chstatus = '';
(async () => {
  const skip = new RegExp(TEAMS_P.SKIP_ROWS);
  const leaves = () => [...document.querySelectorAll('[role="treeitem"]')].filter(e => !e.querySelector('[role="treeitem"]') && !skip.test(e.innerText.trim()));
  const n = leaves().length;
  for (let i = 0; i < n; i++) {
    const el = leaves()[i]; if (!el) break;
    const label = el.innerText.replace(/\s+/g, ' ').trim().slice(0, 80);
    window.__chstatus = i + '/' + n + ' ' + label;
    el.click(); await new Promise(r => setTimeout(r, TEAMS_P.CLICK_WAIT_MS));
    const msgs = (await window.__grab()) || [];
    window.__CH.push({ i, label, title: document.title, msgs: msgs.filter(m => m.mid >= window.__STARTMS) });
  }
  window.__chdone = true;
})();

// Afterwards: summarize to choose the org's chats (authors + counts), then dump the chosen chats as lines
//   "YYYY-MM-DD|HH:MM|teams|<chat>|<author>|<text>"  (local time via new Date(mid)), redacting secrets first,
// into an <article> (see outlook_collect.js dumpOutlook for the pattern) and read it in <=8500-char chunks.
window.summarizeChats = () => JSON.stringify(window.__CH.map(c => c.i + '|' + c.label.slice(0, 24) + '|' + [...new Set(c.msgs.map(m => m.author))].join(',').slice(0, 50) + '|' + c.msgs.length));
