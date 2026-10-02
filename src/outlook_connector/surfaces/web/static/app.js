// Outlook connector local UI. Mail content is untrusted: it is only ever inserted as text.

const TOKEN = document.querySelector('meta[name="session-token"]').content;
const $ = (id) => document.getElementById(id);
// Between people (and file names): display names are often "Last, First", so not a comma.
const SEPARATOR = "; ";

const state = {
  folder: "inbox",       // folder id, "inbox" until the folder list resolves it, or null for the whole mailbox
  folderName: "Inbox",
  mode: "list",          // "list" | "search"
  query: "",
  cursor: null,
  folders: new Map(),    // id -> folder (for well-known aliases)
  threads: new Map(),    // key -> { key, conversationId, messages: Map(id -> summary), expanded, complete, size, sizeAtLeast }
  selectedThreads: new Set(),
  selectedMessages: new Set(),
  activeMessage: null,
  listRequest: 0,        // newest list/search load; older responses are ignored
  readerRequest: 0,      // same for the reader pane
};

// ------------------------------------------------------------------ helpers

function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "class") node.className = value;
    else if (key.startsWith("on")) node.addEventListener(key.slice(2), value);
    else if (value !== undefined && value !== null && value !== false) node.setAttribute(key, value === true ? "" : value);
  }
  for (const child of children.flat()) {
    if (child === null || child === undefined || child === false) continue;
    node.append(child instanceof Node ? child : document.createTextNode(String(child)));
  }
  return node;
}

function showBanner(message) {
  const banner = $("banner");
  banner.textContent = message;
  banner.hidden = !message;
}

async function api(path, options = {}) {
  let response;
  try {
    response = await fetch(path, { ...options, headers: { "X-Session-Token": TOKEN, ...(options.headers || {}) } });
  } catch {
    showBanner("The UI server has stopped. Run `outlook-connector ui` again.");
    throw new Error("server unreachable");
  }
  if (!response.ok) {
    let message = `${response.status} ${response.statusText}`;
    try { message = (await response.json()).error || message; } catch { /* not JSON */ }
    showBanner(message);
    throw new Error(message);
  }
  showBanner("");
  return response;
}

const json = async (path, options) => (await api(path, options)).json();

function fold(text) {
  return (text || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
}

function formatDate(iso) {
  if (!iso) return "";
  const date = new Date(iso);
  const today = new Date();
  return date.toDateString() === today.toDateString()
    ? date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    : date.toLocaleDateString([], { year: "numeric", month: "short", day: "numeric" });
}

function who(person) {
  return person ? (person.name || person.address || "(unknown)") : "(unknown)";
}

function dateBounds() {
  const since = $("since").value ? new Date(`${$("since").value}T00:00:00`).toISOString() : null;
  const until = $("until").value ? new Date(`${$("until").value}T23:59:59.999`).toISOString() : null;
  return { since, until };
}

function query(params) {
  const search = new URLSearchParams();
  for (const [key, value] of Object.entries(params)) if (value !== null && value !== undefined && value !== "") search.set(key, value);
  return search.toString();
}

// ------------------------------------------------------------------ folders

async function loadFolders(refresh = false) {
  if (!$("folders").children.length || refresh) $("folders").replaceChildren(el("li", { class: "loading-item" }, spinner("Loading folders…")));
  const folders = await json(`/api/folders?${query({ refresh })}`);
  state.folders = new Map(folders.map((f) => [f.id, f]));
  const list = $("folders");
  if (state.folder === "inbox") { // the landing folder, requested by alias before its id was known
    const inbox = folders.find((f) => f.well_known === "inbox");
    if (inbox) state.folder = inbox.id;
  }
  list.replaceChildren(folderItem(null, "All mail (recent)", 0, null));
  for (const folder of folders) {
    const depth = (folder.path.match(/\//g) || []).length;
    list.append(folderItem(folder.id, folder.name, depth, folder.unread));
  }
}

function folderItem(id, name, depth, unread) {
  const item = el("li", { class: state.folder === id && state.mode === "list" ? "active" : "", title: name },
    el("span", { style: `padding-left:${depth * 14}px` }, name),
    unread ? el("span", { class: "count" }, unread) : null);
  item.addEventListener("click", () => {
    state.folder = id;
    state.folderName = id ? name : "Recent mail";
    state.mode = "list";
    $("list-title").textContent = state.folderName;
    for (const li of $("folders").children) li.classList.toggle("active", li === item);
    renderExportView();
    loadList(true);
  });
  return item;
}

// ------------------------------------------------------------------ list and search

// Deleted Items, Junk and Sync Issues are left out unless the toggle is on, or the user is inside one
// of them (the server always lists a folder asked for by name; threads, counts and exports follow this).
function includeDeleted() {
  return $("opt-deleted").checked || (state.mode === "list" && insideLeftOutFolder(state.folder));
}

// Deleted Items, Junk Email, Sync Issues, or a folder inside one of them (a folder deleted in
// Outlook moves into Deleted Items), the same rule as the server's.
const LEFT_OUT_FOLDERS = ["deleteditems", "junkemail", "syncissues", "conflicts", "localfailures", "serverfailures"];
function insideLeftOutFolder(id) {
  const seen = new Set();
  for (let folder = state.folders.get(id); folder && !seen.has(folder.id); folder = state.folders.get(folder.parent_id)) {
    seen.add(folder.id);
    if (LEFT_OUT_FOLDERS.includes(folder.well_known)) return true;
  }
  return false;
}

function resetThreads() {
  state.threads = new Map();
  state.cursor = null;
}

function addMessage(summary, { matched = false } = {}) {
  const key = summary.conversation_id || summary.id;
  let thread = state.threads.get(key);
  if (!thread) {
    thread = { key, conversationId: summary.conversation_id, messages: new Map(), expanded: false, complete: false,
      size: undefined, sizeAtLeast: false };
    state.threads.set(key, thread);
  }
  thread.messages.set(summary.id, { ...summary, matched });
}

function spinner(text) {
  return el("div", { class: "loading", role: "status" }, el("span", { class: "spinner", "aria-hidden": "true" }), text);
}

// Every list load gets a number; a response is only applied if no newer load started since.
// This keeps a slow response for a previously clicked folder from replacing the current one.
// previewPath (optional): a local-cache request drawn first, so the list appears instantly while the
// authoritative server response is on its way. The server response then replaces it.
async function loadPage(reset, path, apply, loadingText, previewPath = null) {
  const request = ++state.listRequest;
  if (reset) {
    resetThreads();
    $("threads").replaceChildren(spinner(loadingText));
    $("coverage").textContent = "";
    $("more").hidden = true;
  } else {
    $("more").disabled = true;
    $("more").replaceChildren(el("span", { class: "spinner", "aria-hidden": "true" }), " Loading…");
  }
  if (reset && previewPath) {
    try {
      const cached = await json(previewPath);
      if (request === state.listRequest && cached.items.length) {
        apply(cached);
        render();
        $("more").hidden = true; // the cached cursor is not meaningful
        $("coverage").replaceChildren(el("span", { class: "spinner small", "aria-hidden": "true" }), " cached · updating from Outlook…");
      }
    } catch { /* the server request below reports errors */ }
  }
  try {
    const result = await json(path);
    if (request !== state.listRequest) return; // superseded by a newer click: drop this response
    if (reset) resetThreads(); // drop the cached preview rows
    apply(result);
    state.cursor = result.cursor;
    showCoverage(result.coverage);
    render();
    loadSizes(request);
  } catch {
    if (request !== state.listRequest) return;
    if (reset) $("threads").replaceChildren(el("p", { class: "muted pad" }, "Could not load messages (see the message above)."));
  } finally {
    if (request === state.listRequest) {
      $("more").disabled = false;
      $("more").textContent = "Load older messages";
    }
  }
}

function loadList(reset) {
  const { since, until } = dateBounds();
  const scope = { folder: state.folder, since, until, limit: 100, include_deleted_items: includeDeleted() };
  const path = `/api/messages?${query({ ...scope, cursor: reset ? null : state.cursor })}`;
  const preview = reset ? `/api/messages?${query({ ...scope, refresh: false })}` : null;
  return loadPage(reset, path, (page) => {
    for (const item of page.items) addMessage(item);
  }, "Loading messages…", preview);
}

function runSearch(reset) {
  const { since, until } = dateBounds();
  const path = `/api/search?${query({ q: state.query, since, until, folder: state.folder, limit: 50,
    include_deleted_items: includeDeleted(), cursor: reset ? null : state.cursor })}`;
  return loadPage(reset, path, (result) => {
    for (const hit of result.conversations) for (const message of hit.matching_messages) addMessage(message, { matched: true });
  }, "Searching the mailbox…");
}

// Ask Outlook how many messages each listed conversation really has, so single messages render as
// plain rows and threads show an accurate count. Rows look as before until the counts arrive.
async function loadSizes(request) {
  const pending = [...state.threads.values()].filter((t) => t.conversationId && t.size === undefined).map((t) => t.conversationId);
  for (let start = 0; start < pending.length; start += 200) {
    let sizes;
    try {
      sizes = await json("/api/thread-sizes", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ conversation_ids: pending.slice(start, start + 200), include_deleted_items: includeDeleted() }) });
    } catch {
      return; // counts are a refinement; the rows keep working without them
    }
    if (request !== state.listRequest) return; // the list was replaced meanwhile
    const byId = new Map(sizes.map((s) => [s.conversation_id, s]));
    for (const thread of state.threads.values()) {
      const size = byId.get(thread.conversationId);
      if (size) { thread.size = size.messages; thread.sizeAtLeast = size.at_least; }
    }
    render();
  }
}

function showCoverage(coverage) {
  const parts = [];
  const excluded = coverage.excluded || {};
  const notShown = (excluded.deleted_or_junk || 0) + (excluded.sync_issues || 0);
  if (notShown) parts.push(`${notShown} in Deleted / Junk / Sync Issues not shown`);
  $("coverage").textContent = parts.join(" · ");
  $("coverage").title = (coverage.notes || []).join("\n");
}

// ------------------------------------------------------------------ rendering

function sortedThreads() {
  const threads = [...state.threads.values()];
  if (state.mode === "search") return threads; // server rank order
  const latest = (t) => Math.max(...[...t.messages.values()].map((m) => Date.parse(m.received_at || m.sent_at || 0)));
  return threads.sort((a, b) => latest(b) - latest(a));
}

function matchesFilter(thread, needle) {
  if (!needle) return true;
  for (const m of thread.messages.values()) {
    const haystack = [m.subject, m.preview, m.folder, m.sender && m.sender.name, m.sender && m.sender.address,
      ...(m.to || []).map((r) => `${r.name} ${r.address}`)].join(" ");
    if (fold(haystack).includes(needle)) return true;
  }
  return false;
}

function render() {
  const needle = fold($("filter").value.trim());
  const container = $("threads");
  container.replaceChildren(...sortedThreads().filter((t) => matchesFilter(t, needle)).map(renderThread));
  if (!container.children.length) container.append(el("p", { class: "muted", style: "padding:0 12px" }, "No messages."));
  $("more").hidden = !state.cursor;
  renderSelection();
}

function isSingle(thread) {
  if (thread.messages.size !== 1) return false;
  if (!thread.conversationId) return true;
  return thread.size !== undefined && !thread.sizeAtLeast && thread.size <= 1;
}

function countLabel(thread, loaded) {
  if (thread.complete) return `${loaded}`;
  if (thread.size !== undefined) return `${Math.max(thread.size, loaded)}${thread.sizeAtLeast ? "+" : ""}`;
  return `${loaded}+`;
}

// A conversation with one message: a plain row, like Outlook's conversation view.
function renderSingle(thread) {
  const [message] = thread.messages.values();
  const checkbox = el("input", { type: "checkbox", title: "Export this message",
    onclick: (event) => event.stopPropagation(),
    onchange: (event) => {
      toggle(state.selectedMessages, message.id, event.target.checked);
      if (!event.target.checked && thread.conversationId) state.selectedThreads.delete(thread.conversationId);
      render();
    } });
  checkbox.checked = state.selectedMessages.has(message.id) || state.selectedThreads.has(thread.conversationId);
  return el("div", { class: "thread" },
    el("div", { class: `thread-row${message.is_read === false ? " unread" : ""}${state.activeMessage === message.id ? " active" : ""}`,
      onclick: () => openMessage(message.id) },
      checkbox,
      el("span", { class: "toggle" }),
      el("div", {},
        el("div", { class: "subject" }, message.subject || "(no subject)", ...badges(message, state.folder === null)),
        el("div", { class: "who" }, who(message.sender))),
      el("span", { class: "date" }, formatDate(message.received_at || message.sent_at))));
}

function renderThread(thread) {
  if (isSingle(thread)) return renderSingle(thread);
  // newest on top, like the list and Outlook's conversation view (exports and get_thread stay oldest first)
  const when = (m) => Date.parse(m.received_at || m.sent_at || 0);
  const messages = [...thread.messages.values()].sort((a, b) => when(b) - when(a));
  const newest = messages[0];
  const senders = [...new Set(messages.map((m) => who(m.sender)))].join(SEPARATOR);
  const unread = messages.some((m) => m.is_read === false);
  const selectable = Boolean(thread.conversationId);
  const checkbox = el("input", { type: "checkbox", title: "Export the whole thread", disabled: !selectable,
    onclick: (event) => event.stopPropagation(),
    onchange: (event) => { toggle(state.selectedThreads, thread.conversationId, event.target.checked); render(); } });
  checkbox.checked = state.selectedThreads.has(thread.conversationId);
  const row = el("div", { class: `thread-row${unread ? " unread" : ""}`, onclick: () => expand(thread) },
    checkbox,
    el("span", { class: "toggle" }, thread.expanded ? "▾" : "▸"),
    el("div", {},
      el("div", { class: "subject" }, newest.subject || "(no subject)",
        el("span", { class: "badge" }, countLabel(thread, messages.length))),
      el("div", { class: "who" }, senders)),
    el("span", { class: "date" }, formatDate(newest.received_at || newest.sent_at)));
  const node = el("div", { class: "thread" }, row);
  if (thread.expanded) {
    node.append(el("div", { class: "messages" }, messages.map((m) => renderMessage(m, thread)),
      thread.loading ? spinner("Loading the whole conversation…") : null));
  }
  return node;
}

// Folder (optional) and "also in" for merged copies of one message.
function badges(message, withFolder) {
  return [
    withFolder && message.folder ? el("span", { class: "badge" }, message.folder) : null,
    ...(message.also_in || []).map((folder) => el("span", { class: "badge copy", title: "Another copy of this message" }, `also in ${folder}`)),
  ];
}

function renderMessage(message, thread) {
  const covered = state.selectedThreads.has(thread.conversationId);
  const checkbox = el("input", { type: "checkbox", disabled: covered, title: covered ? "Included with the thread" : "Export this message",
    onclick: (event) => event.stopPropagation(),
    onchange: (event) => { toggle(state.selectedMessages, message.id, event.target.checked); renderSelection(); } });
  checkbox.checked = covered || state.selectedMessages.has(message.id);
  const classes = ["msg-row", state.activeMessage === message.id && "active", message.is_read === false && "unread",
    message.matched && "matched"].filter(Boolean).join(" ");
  return el("div", { class: classes, title: message.matched ? "Matches the search" : undefined, onclick: () => openMessage(message.id) },
    checkbox,
    el("div", {},
      el("div", { class: "msg-subject" }, who(message.sender), ...badges(message, true)),
      el("div", { class: "who" }, message.preview || message.subject || "")),
    el("span", { class: "date" }, formatDate(message.received_at || message.sent_at)));
}

function toggle(set, value, on) {
  if (on) set.add(value); else set.delete(value);
}

async function expand(thread) {
  thread.expanded = !thread.expanded;
  if (thread.expanded && thread.conversationId && !thread.complete && !thread.loading) {
    thread.loading = true;
    render();
    try {
      const full = await json(`/api/threads/${encodeURIComponent(thread.conversationId)}?${query({ include_deleted_items: includeDeleted() })}`);
      for (const entry of full.messages) thread.messages.set(entry.message.id, { ...entry.message, matched: thread.messages.get(entry.message.id)?.matched });
      // incomplete coverage: the conversation is larger than the server lists (a "1000+" thread)
      thread.complete = full.coverage.complete;
      thread.size = full.messages.length;
      thread.sizeAtLeast = !full.coverage.complete;
    } finally {
      thread.loading = false;
    }
  }
  if (state.threads.get(thread.key) === thread) render(); // skip if the list was replaced meanwhile
}

// ------------------------------------------------------------------ reader

async function openMessage(id) {
  state.activeMessage = id;
  render();
  const request = ++state.readerRequest;
  $("reader-title").replaceChildren(spinner("Loading message…"));
  $("reader-meta").replaceChildren();
  $("reader-body").textContent = "";
  const body = $("reader-full").checked ? "full" : "unique";
  let content;
  try {
    content = await json(`/api/messages/${encodeURIComponent(id)}?${query({ body })}`);
  } catch {
    if (request === state.readerRequest) $("reader-title").textContent = "Could not load this message.";
    return;
  }
  if (request !== state.readerRequest) return; // another message was clicked meanwhile
  const m = content.message;
  $("reader-title").textContent = m.subject || "(no subject)";
  const rows = [["From", who(m.sender)], ["To", (m.to || []).map(who).join(SEPARATOR)], ["Cc", (m.cc || []).map(who).join(SEPARATOR)],
    ["Date", m.received_at ? new Date(m.received_at).toLocaleString() : ""],
    ["Folder", [m.folder, ...(m.also_in || [])].filter(Boolean).join(SEPARATOR)]];
  const files = (content.attachments || []).filter((a) => !a.is_inline);
  if (files.length) rows.push(["Attachments", files.map((a) => attachmentButton(m.id, a))]);
  $("reader-meta").replaceChildren(...rows.filter(([, v]) => v && v.length).flatMap(([k, v]) => [el("dt", {}, k), el("dd", {}, v)]));
  $("reader-body").textContent = content.text;
}

function attachmentButton(messageId, attachment) {
  const label = attachment.name || "attachment";
  if (attachment.kind === "reference") return el("span", { class: "muted attachment" }, `${label} (cloud link)`);
  return el("button", { type: "button", class: "attachment", title: "Download",
    onclick: () => download(`/api/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachment.id)}`, {}) }, label);
}

// Fetch with the session token and hand the response to the browser as a download.
async function download(path, options) {
  const response = await api(path, options);
  const url = URL.createObjectURL(await response.blob());
  const link = el("a", { href: url, download: filenameFrom(response.headers.get("Content-Disposition")) });
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

// ------------------------------------------------------------------ selection and export

function renderSelection() {
  const threads = state.selectedThreads.size;
  const messages = [...state.selectedMessages].filter((id) => !coveredByThread(id)).length;
  const parts = [];
  if (threads) parts.push(`${threads} thread${threads > 1 ? "s" : ""}`);
  if (messages) parts.push(`${messages} message${messages > 1 ? "s" : ""}`);
  $("selection").textContent = parts.length ? `Selected: ${parts.join(" + ")}` : "Nothing selected";
  $("export").disabled = !parts.length;
}

function coveredByThread(messageId) {
  for (const thread of state.threads.values()) {
    if (thread.messages.has(messageId) && state.selectedThreads.has(thread.conversationId)) return true;
  }
  return false;
}

function exportOptions() {
  return {
    include_attachments: $("opt-attachments").checked,
    combine: $("opt-files").value,
    body: $("opt-full").checked ? "full" : "unique",
    include_deleted_items: includeDeleted(),
  };
}

function exportRequest() {
  return {
    conversation_ids: [...state.selectedThreads],
    message_ids: [...state.selectedMessages].filter((id) => !coveredByThread(id)),
    ...exportOptions(),
  };
}

// The whole current view: this folder (or the mailbox) within the chosen dates, up to 2,000 messages.
function viewRequest() {
  const { since, until } = dateBounds();
  return { folder: state.folder, since, until, ...exportOptions() };
}

function renderExportView() {
  const button = $("export-view");
  const { since, until } = dateBounds();
  button.hidden = state.mode !== "list";
  button.disabled = !state.folder && !since && !until; // the whole mailbox needs a date range
  button.title = button.disabled ? "Choose a date range to export from all mail"
    : "Export every message of this folder and date range (up to 2,000)";
}

function filenameFrom(disposition) {
  const star = /filename\*=utf-8''([^;]+)/i.exec(disposition || "");
  if (star) return decodeURIComponent(star[1]);
  const plain = /filename="?([^";]+)"?/i.exec(disposition || "");
  return plain ? plain[1] : "outlook-export";
}

async function runExport(button, request, label) {
  button.disabled = true;
  button.textContent = "Exporting…";
  try {
    await download("/api/export", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(request) });
  } catch { /* the banner shows the error */ } finally {
    button.textContent = label;
    renderSelection();
    renderExportView();
  }
}

// ------------------------------------------------------------------ wiring

$("search-form").addEventListener("submit", (event) => {
  event.preventDefault();
  state.query = $("search").value.trim();
  state.mode = state.query ? "search" : "list";
  $("list-title").textContent = state.query ? `Search: ${state.query}` : state.folderName;
  renderExportView();
  (state.query ? runSearch : loadList)(true);
});
$("filter").addEventListener("input", render);

// ------------------------------------------------------------------ date range picker (same behaviour as lrh-teams)

const picker = { month: new Date(), choosingEnd: false, draftStart: null, applied: "|" }; // applied = "since|until"
picker.month.setDate(1);

function dateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function dateFromKey(value) {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function formatKey(value) {
  return dateFromKey(value).toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}

function setDateRange(start, end) {
  $("since").value = start || "";
  $("until").value = end || "";
  $("range-label").textContent = !start ? "Any date"
    : start === end ? formatKey(start) : `${formatKey(start)} – ${formatKey(end)}`;
  renderCalendar();
}

function setCalendarOpen(open) {
  $("calendar").hidden = !open;
  $("range-toggle").setAttribute("aria-expanded", String(open));
  if (open) {
    // open towards whichever side has room for the 330px calendar
    const box = $("range-picker").getBoundingClientRect();
    $("calendar").classList.toggle("align-right", box.left + 340 > window.innerWidth);
    picker.month = $("since").value ? dateFromKey($("since").value) : new Date();
    picker.month.setDate(1);
    picker.choosingEnd = false;
    picker.draftStart = null;
    renderCalendar();
    return;
  }
  const range = `${$("since").value}|${$("until").value}`;
  if (range !== picker.applied) { // reload only when the range actually changed
    picker.applied = range;
    renderExportView();
    (state.mode === "search" ? runSearch : loadList)(true);
  }
}

function renderCalendar() {
  $("calendar-month").textContent = picker.month.toLocaleDateString(undefined, { month: "long", year: "numeric" });
  $("calendar-hint").textContent = picker.choosingEnd
    ? "Choose an end date, or press Done to keep one day."
    : "Click once for one day; click again to set the end date.";
  const first = new Date(picker.month.getFullYear(), picker.month.getMonth(), 1);
  const mondayOffset = (first.getDay() + 6) % 7;
  const start = $("since").value;
  const end = $("until").value;
  const days = [];
  for (let index = 0; index < 42; index += 1) {
    const day = new Date(first.getFullYear(), first.getMonth(), 1 - mondayOffset + index);
    const key = dateKey(day);
    const classes = ["calendar-day"];
    if (day.getMonth() !== picker.month.getMonth()) classes.push("outside-month");
    if (start && end && key >= start && key <= end) classes.push("in-range");
    if (start && key === start) classes.push("range-start");
    if (end && key === end) classes.push("range-end");
    days.push(el("button", { type: "button", class: classes.join(" "), "aria-label": day.toLocaleDateString(undefined, { dateStyle: "full" }),
      onclick: () => chooseDate(key) }, day.getDate()));
  }
  $("calendar-days").replaceChildren(...days);
}

function chooseDate(key) {
  if (!picker.choosingEnd || key < picker.draftStart) {
    picker.draftStart = key;
    picker.choosingEnd = true;
    setDateRange(key, key);
    return;
  }
  setDateRange(picker.draftStart, key);
  picker.choosingEnd = false;
  setCalendarOpen(false);
}

$("range-toggle").addEventListener("click", () => setCalendarOpen($("calendar").hidden));
$("previous-month").addEventListener("click", () => { picker.month.setMonth(picker.month.getMonth() - 1); renderCalendar(); });
$("next-month").addEventListener("click", () => { picker.month.setMonth(picker.month.getMonth() + 1); renderCalendar(); });
$("clear-dates").addEventListener("click", () => { picker.choosingEnd = false; setDateRange("", ""); setCalendarOpen(false); });
$("calendar-done").addEventListener("click", () => setCalendarOpen(false));
document.addEventListener("pointerdown", (event) => {
  if (!$("calendar").hidden && !$("range-picker").contains(event.target)) setCalendarOpen(false);
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !$("calendar").hidden) setCalendarOpen(false);
});
$("more").addEventListener("click", () => (state.mode === "search" ? runSearch : loadList)(false));
$("export-view").addEventListener("click", () => runExport($("export-view"), viewRequest(), "export this view"));
$("refresh-folders").addEventListener("click", () => loadFolders(true));
$("reader-full").addEventListener("change", () => state.activeMessage && openMessage(state.activeMessage));
$("opt-deleted").addEventListener("change", () => (state.mode === "search" ? runSearch : loadList)(true));
$("clear").addEventListener("click", () => { state.selectedThreads.clear(); state.selectedMessages.clear(); render(); });
$("export").addEventListener("click", () => runExport($("export"), exportRequest(), "Export"));
setInterval(() => api("/api/heartbeat", { method: "POST" }).catch(() => {}), 60_000);

(async function start() {
  const status = await json("/api/status");
  $("account").textContent = status.account || "";
  if (!status.signed_in.read) {
    showBanner(`Not signed in. Run \`${status.sign_in_command}\` in a terminal, then reload this page.`);
    return;
  }
  renderExportView();
  await Promise.all([loadFolders(), loadList(true)]); // independent: load side by side
})();
