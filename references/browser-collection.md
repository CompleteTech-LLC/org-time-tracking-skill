# Browser collection: Outlook, Teams and calendar mechanics

The collectors in `scripts/*.js` run inside the operator's signed-in browser tab through the agent's JavaScript tool. Their only inputs are the `PARAMS` block at the top of each script or `window.__PARAMS`; each requires `START` and refuses to run without it.

## Getting data out of the page (the JS tool has a tiny output cap)
The JavaScript tool returns only about 800 characters and blocks anything that looks like a query string or base64. Transfer large data like this:
1. Collect into a `window.*` variable (data persists until the page navigates; **navigating reloads the app and wipes it**).
2. Create an `<article>` element fixed over the page (`position:fixed;inset:0;z-index:99999;background:#fff;white-space:pre-wrap`) and set its `textContent` to at most about 8,500 characters of lines.
3. Call the browser tool's *get page text* action; it returns the article text. Repeat per chunk.
4. Remove the element afterwards.

Long loops: a single JS call times out after about 45 s. Start the loop as a background async function, return immediately, then poll a `window.__done` flag with short waits. Scripts use distinct global names (`OUTLOOK_P`, `TEAMS_P`, `MEET_P`, `CAL_P`) so several can be pasted into one page.

Posting data from the page to `http://localhost` is blocked by the site's content-security policy; do not work around it. Never put page data in URLs or send it to other hosts.

## Outlook web (https://outlook.office.com/mail/)
Run two searches in the search box, in this order:
1. `from:<domain>`: everything the org sent. Split it into `staff/` (human senders) and `alerts/` (system senders: monitoring or helpdesk mailboxes, "Warning:", "Request ID", "dropped", "configured"; calendar auto-responses can go to staff). Judge by sender display name and subject; if unsure, ask or note it in the README.
2. The domain name without its TLD (broad search): keep results whose label is not already in the first set. That is `related/`.

Collecting the list (`scripts/outlook_collect.js`):
- Rows are `div[data-convid]` with an `aria-label` holding "[Unread] Sender Subject Date Preview". The scroll container is `ANCESTOR_LEVELS` parents above the first row. Scroll in steps, dedupe by label prefix, stop once the dates pass the start date or the count stops growing.
- Dates in labels are `M/D/YYYY` for older mail and weekday-only for the last week; resolve weekday labels against today's date. List labels carry no time for older mail, so times are not recorded.
- Verify the search by opening one result and checking the sender address in the reading pane.
- A month with zero results may be a search limit; say so in the README rather than asserting there was no mail.

## Teams web (https://teams.microsoft.com/v2/, redirects to teams.cloud.microsoft)
Wait for the "setting things up" screen to finish (about 20 s). Then (`scripts/teams_collect.js`):
- Chat rows are leaf `[role="treeitem"]` elements. Click "See more" first.
- Messages: `[data-tid="chat-pane-item"]` inside `[data-tid="message-pane-list-viewport"]`; the inner `[data-tid="chat-pane-message"]` has `data-mid` = epoch milliseconds. The author is `[data-tid="message-author-name"]`; consecutive messages from one author omit it (carry forward). Own messages contain `ME_MARK` and have no name.
- Scroll the viewport to the top repeatedly until the oldest `data-mid` is before the start date; the pane lazy-loads.
- The page title lags after a click. Verify the capture by checking author names and the first message, not by the title alone.
- Convert `data-mid` to the operator's time zone before grouping by day.
- Decide which chats belong to the org by participants (the header shows the email for 1:1 external chats, group chats list names), by who posts, and by content. Exclude unrelated chats (clubs, personal) and say which and why. Say that channels and communities were not checked unless they were.
- Many chats will have no messages since the start date; that is normal.

## Calendar and meeting cards
1. **Outlook calendar** (`https://outlook.office.com/calendar/view/month`; `scripts/calendar_collect.js`). Check the "Filter applied" menu first and note hidden types (for example Reservations). Event buttons carry an aria-label like `Title, 1:00 PM to 1:30 PM, Monday, September 28, 2026, By Organizer, Busy, Recurring event`. **Cancelled occurrences vanish from the calendar**; find cancellations from "Meeting canceled" emails and chat remarks and list them as inferred. Tag "Exception to recurring event" rows (moved instances).
2. **Teams meeting cards** (`scripts/teams_meetings.js`): each meeting chat logs "Meeting started at ..." and "Meeting ended: at ... after N minutes S seconds". These are the only real durations. Sweep every chat; attribute events by `document.title` after the scan, not by the row clicked. A chat that logged every call historically but only a few recently means the missing calls probably did not happen.
3. Recaps and transcripts live in each meeting chat's Recap tab and are not reachable by these scripts; say so unless you opened them.
