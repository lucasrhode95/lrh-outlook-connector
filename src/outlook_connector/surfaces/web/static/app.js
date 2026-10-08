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
  conversations: new Map(),    // key -> { key, conversationId, messages: Map(id -> summary), expanded, complete, size, sizeAtLeast }
  selectedConversations: new Set(),
  selectedMessages: new Set(),
  attachments: new Map(), // message id -> its files, for the list's chips
  attachmentRequests: new Set(), // message ids currently being enriched with file names
  activeMessage: null,
  listRequest: 0,        // newest list/search load; older responses are ignored
  folderRequest: 0,      // newest folder load; older refreshes are ignored
  readerRequest: 0,      // same for the reader pane
};

let bannerOwner = null;

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

function showBanner(message, owner = null) {
  const banner = $("banner");
  banner.textContent = message;
  banner.hidden = !message;
  bannerOwner = message ? owner : null;
}

// A retry clears only its own previous error; unrelated background successes leave it visible.
function clearBanner(owner) {
  if (owner && bannerOwner === owner) showBanner("");
}

async function api(path, options = {}, owner = null, isCurrent = () => true) {
  let response;
  try {
    response = await fetch(path, { ...options, headers: { "X-Session-Token": TOKEN, ...(options.headers || {}) } });
  } catch {
    if (isCurrent()) showBanner("The UI server has stopped. Run `outlook-connector ui` again.", owner);
    throw new Error("server unreachable");
  }
  if (!response.ok) {
    let message = `${response.status} ${response.statusText}`;
    try { message = (await response.json()).error || message; } catch { /* not JSON */ }
    if (isCurrent()) showBanner(message, owner);
    throw new Error(message);
  }
  return response;
}

const json = async (path, options, owner, isCurrent) => (await api(path, options, owner, isCurrent)).json();

function fold(text) {
  return (text || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
}

function bindClearableSearch(input, clearButton, { onInput = () => {}, onClear = onInput, showClearWhenEmpty = false } = {}) {
  const updateClearButton = () => { clearButton.hidden = !showClearWhenEmpty && !input.value; };
  input.addEventListener("input", () => {
    updateClearButton();
    onInput();
  });
  clearButton.addEventListener("click", () => {
    input.value = "";
    updateClearButton();
    onClear();
  });
  updateClearButton();
}

function bindExpandableSearch({ host, toggle, input, clearButton, onInput, hideTitle = false }) {
  function close() {
    const wasOpen = host.classList.contains("search-open");
    input.value = "";
    host.classList.remove("search-open");
    toggle.setAttribute("aria-expanded", "false");
    clearButton.hidden = true;
    if (wasOpen) onInput();
    toggle.focus();
  }

  function open() {
    host.classList.add("search-open");
    toggle.setAttribute("aria-expanded", "true");
    clearButton.hidden = false;
    requestAnimationFrame(() => input.focus());
  }

  toggle.addEventListener("click", () => {
    if (host.classList.contains("search-open")) close();
    else open();
  });
  bindClearableSearch(input, clearButton, { onInput, onClear: close, showClearWhenEmpty: true });
  input.addEventListener("keydown", (event) => {
    if (event.key === "Escape") close();
  });
  if (hideTitle) host.classList.add("hide-search-title");
}

// Like Outlook: today "9:31 AM", this week "Fri 9:31 AM", this year "Sep 28", older "Sep 28, 2025".
function formatDate(iso) {
  if (!iso) return "";
  const date = new Date(iso);
  const now = new Date();
  const time = date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  if (date.toDateString() === now.toDateString()) return time;
  if (now - date < 6 * 86_400_000 && date < now) return `${date.toLocaleDateString([], { weekday: "short" })} ${time}`;
  const options = { month: "short", day: "numeric" };
  if (date.getFullYear() !== now.getFullYear()) options.year = "numeric";
  return date.toLocaleDateString([], options);
}

// ------------------------------------------------------------------ icons (inline: the UI loads nothing from the internet)

const ICONS = {
  chevronRight: '<path d="M9 6l6 6l-6 6"/>',
  chevronDown: '<path d="M6 9l6 6l6 -6"/>',
  paperclip: '<path d="M15 7l-6.5 6.5a1.5 1.5 0 0 0 3 3l6.5 -6.5a3 3 0 0 0 -6 -6l-6.5 6.5a4.5 4.5 0 0 0 9 9l6.5 -6.5"/>',
  flag: '<path d="M5 5a5 5 0 0 1 7 0a5 5 0 0 0 7 0v9a5 5 0 0 1 -7 0a5 5 0 0 0 -7 0v-9z"/><path d="M5 21v-7"/>',
  calendar: '<path d="M4 7a2 2 0 0 1 2 -2h12a2 2 0 0 1 2 2v12a2 2 0 0 1 -2 2h-12a2 2 0 0 1 -2 -2z"/><path d="M16 3v4"/><path d="M8 3v4"/><path d="M4 11h16"/>',
  calendarOff: '<path d="M4 7a2 2 0 0 1 2 -2h12a2 2 0 0 1 2 2v12a2 2 0 0 1 -2 2h-12a2 2 0 0 1 -2 -2z"/><path d="M16 3v4"/><path d="M8 3v4"/><path d="M4 11h16"/><path d="M10 14l4 4m0 -4l-4 4"/>',
  file: '<path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21h-10a2 2 0 0 1 -2 -2v-14a2 2 0 0 1 2 -2h7l5 5v11a2 2 0 0 1 -2 2z"/>',
};

function icon(name, extraClass = "") {
  const span = el("span", { class: `icon ${extraClass}`.trim(), "aria-hidden": "true" });
  span.innerHTML = `<svg viewBox="0 0 24 24">${ICONS[name]}</svg>`; // fixed markup above, never mail content
  return span;
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
  const request = ++state.folderRequest;
  clearBanner("folders");
  if (!$("folders").children.length || refresh) $("folders").replaceChildren(el("li", { class: "loading-item" }, spinner("Loading folders…")));
  const list = $("folders");
  try {
    const folders = await json(`/api/folders?${query({ refresh })}`, undefined, "folders", () => request === state.folderRequest);
    if (request !== state.folderRequest) return;
    state.folders = new Map(folders.map((f) => [f.id, f]));
    if (state.folder === "inbox") { // the landing folder, requested by alias before its id was known
      const inbox = folders.find((f) => f.well_known === "inbox");
      if (inbox) state.folder = inbox.id;
    }
    list.replaceChildren(folderItem(null, "All mail (recent)", 0, null));
    for (const folder of folders) {
      const depth = (folder.path.match(/\//g) || []).length;
      list.append(folderItem(folder.id, folder.name, depth, folder.unread));
    }
    list.append(el("li", { id: "folder-search-empty", class: "folder-empty", hidden: true }, "No folders match that search."));
    renderFolderFilter();
  } catch {
    if (request === state.folderRequest) {
      list.replaceChildren(el("li", { class: "folder-load-error", role: "status" }, "Could not load folders. Use refresh to try again."));
    }
  }
}

function renderFolderFilter() {
  const needle = fold($("folder-filter").value.trim());
  const items = [...$("folders").querySelectorAll("[data-folder-name]")];
  let visible = 0;
  for (const item of items) {
    const matches = fold(item.dataset.folderName).includes(needle);
    item.hidden = !matches;
    if (matches) visible += 1;
  }
  const empty = $("folder-search-empty");
  if (empty) empty.hidden = !needle || visible > 0;
}

function folderItem(id, name, depth, unread) {
  const active = state.folder === id && state.mode === "list";
  const button = el("button", { type: "button", class: "folder-button", title: name, "aria-current": active ? "page" : undefined },
    el("span", { class: "folder-name", style: `padding-left:${depth * 14}px` }, name),
    unread ? el("span", { class: "count" }, unread) : null);
  const item = el("li", { class: active ? "active" : "", "data-folder-name": name }, button);
  button.addEventListener("click", () => {
    state.folder = id;
    state.folderName = id ? name : "Recent mail";
    state.mode = "list";
    $("list-title").textContent = state.folderName;
    for (const li of $("folders").children) {
      const selected = li === item;
      li.classList.toggle("active", selected);
      const folderButton = li.querySelector(".folder-button");
      if (selected) folderButton?.setAttribute("aria-current", "page");
      else folderButton?.removeAttribute("aria-current");
    }
    renderExportView();
    loadList(true);
  });
  return item;
}

// ------------------------------------------------------------------ list and search

// Deleted Items and Junk are left out unless the toggle is on, or the user is inside one
// of them (the server always lists a folder asked for by name; conversations, counts and exports follow this).
function showDeletedItems() {
  return $("opt-deleted").checked || (state.mode === "list" && insideLeftOutFolder(state.folder));
}

// Deleted Items, Junk Email, or a folder inside one of them (a folder deleted in Outlook moves into
// Deleted Items), the same rule as the server's.
const LEFT_OUT_FOLDERS = ["deleteditems", "junkemail"];
function insideLeftOutFolder(id) {
  const seen = new Set();
  for (let folder = state.folders.get(id); folder && !seen.has(folder.id); folder = state.folders.get(folder.parent_id)) {
    seen.add(folder.id);
    if (LEFT_OUT_FOLDERS.includes(folder.well_known)) return true;
  }
  return false;
}

function resetConversations() {
  state.conversations = new Map();
  state.cursor = null;
}

function addMessage(summary, { matched = false } = {}) {
  const key = summary.conversation_id || summary.id;
  let conversation = state.conversations.get(key);
  if (!conversation) {
    conversation = { key, conversationId: summary.conversation_id, messages: new Map(), expanded: false, complete: false,
      size: undefined, sizeAtLeast: false };
    state.conversations.set(key, conversation);
  }
  conversation.messages.set(summary.id, { ...summary, matched });
}

function spinner(text) {
  return el("div", { class: "loading", role: "status" }, el("span", { class: "spinner", "aria-hidden": "true" }), text);
}

// Every list load gets a number; a response is only applied if no newer load started since.
// This keeps a slow response for a previously clicked folder from replacing the current one.
async function loadPage(reset, path, apply, loadingText) {
  const request = ++state.listRequest;
  clearBanner("list");
  if (reset) {
    resetConversations();
    $("conversations").replaceChildren(spinner(loadingText));
    $("coverage").textContent = "";
    $("more").hidden = true;
  } else {
    $("more").disabled = true;
    $("more").replaceChildren(el("span", { class: "spinner", "aria-hidden": "true" }), " Loading…");
  }
  try {
    const result = await json(path, undefined, "list", () => request === state.listRequest);
    if (request !== state.listRequest) return; // superseded by a newer click: drop this response
    apply(result);
    state.cursor = result.cursor;
    showCoverage(result.coverage);
    render();
    loadSizes(request);
    loadAttachmentNames();
  } catch {
    if (request !== state.listRequest) return;
    if (reset) $("conversations").replaceChildren(el("p", { class: "muted pad" }, "Could not load messages (see the message above)."));
  } finally {
    if (request === state.listRequest) {
      $("more").disabled = false;
      $("more").textContent = "Load older messages";
    }
  }
}

function loadList(reset) {
  const { since, until } = dateBounds();
  const scope = { folder: state.folder, since, until, limit: 100, deleted_items: showDeletedItems(),
    meeting_mail: $("opt-meetings").checked };
  const path = `/api/messages?${query({ ...scope, cursor: reset ? null : state.cursor })}`;
  return loadPage(reset, path, (page) => {
    for (const item of page.items) addMessage(item);
  }, "Loading messages…");
}

function runSearch(reset) {
  const { since, until } = dateBounds();
  const path = `/api/search?${query({ q: state.query, since, until, folder: state.folder, limit: 50,
    deleted_items: showDeletedItems(), meeting_mail: $("opt-meetings").checked,
    cursor: reset ? null : state.cursor })}`;
  return loadPage(reset, path, (result) => {
    for (const hit of result.conversations) for (const message of hit.matching_messages) addMessage(message, { matched: true });
  }, "Searching the mailbox…");
}

// Ask Outlook how many messages each listed conversation really has, so single messages render as
// plain rows and conversations show an accurate count. Rows look as before until the counts arrive.
async function loadSizes(request) {
  const pending = [...state.conversations.values()].filter((t) => t.conversationId && t.size === undefined).map((t) => t.conversationId);
  if (pending.length) clearBanner("sizes");
  for (let start = 0; start < pending.length; start += 200) {
    let sizes;
    try {
      sizes = await json("/api/conversation-sizes", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ conversation_ids: pending.slice(start, start + 200), scope: { deleted_items: showDeletedItems() } }) },
        "sizes", () => request === state.listRequest);
    } catch {
      return; // counts are a refinement; the rows keep working without them
    }
    if (request !== state.listRequest) return; // the list was replaced meanwhile
    const byId = new Map(sizes.map((s) => [s.conversation_id, s]));
    for (const conversation of state.conversations.values()) {
      const size = byId.get(conversation.conversationId);
      if (size) { conversation.size = size.messages; conversation.sizeAtLeast = size.at_least; }
    }
    render();
  }
}

// File names for the chips under messages with attachments: one batched request per 200 messages
// (Graph takes them 20 at a time). Rows work without them.
async function loadAttachmentNames() {
  const pending = [...state.conversations.values()].flatMap((t) => [...t.messages.values()])
    .filter((m) => m.has_attachments && !state.attachments.has(m.id) && !state.attachmentRequests.has(m.id)).map((m) => m.id);
  if (pending.length) clearBanner("attachments");
  for (let start = 0; start < pending.length; start += 200) {
    const ids = pending.slice(start, start + 200);
    for (const id of ids) state.attachmentRequests.add(id);
    let names;
    const current = () => ids.some((id) => [...state.conversations.values()].some((conversation) => conversation.messages.has(id)));
    try {
      names = await json("/api/attachments", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message_ids: ids }) }, "attachments", current);
    } catch {
      for (const id of ids) state.attachmentRequests.delete(id);
      return;
    }
    for (const id of ids) state.attachmentRequests.delete(id);
    for (const [id, items] of Object.entries(names)) state.attachments.set(id, items);
    const visibleIds = new Set([...state.conversations.values()].flatMap((conversation) => [...conversation.messages.keys()]));
    if (Object.keys(names).some((id) => visibleIds.has(id))) render();
  }
}

function showCoverage(coverage) {
  const parts = [];
  const excluded = coverage.excluded || {};
  if (excluded.deleted_or_junk) parts.push(`${excluded.deleted_or_junk} in Deleted / Junk not shown`);
  if (excluded.meeting_mail) parts.push(`${excluded.meeting_mail} meeting messages not shown`);
  $("coverage").textContent = parts.join(" · ");
  $("coverage").title = (coverage.notes || []).join("\n");
}

// ------------------------------------------------------------------ rendering

function sortedConversations() {
  const conversations = [...state.conversations.values()];
  if (state.mode === "search") return conversations; // server rank order
  const latest = (t) => Math.max(...[...t.messages.values()].map((m) => Date.parse(m.received_at || 0)));
  return conversations.sort((a, b) => latest(b) - latest(a));
}

function matchesFilter(conversation, needle) {
  if (!needle) return true;
  for (const m of conversation.messages.values()) {
    const haystack = [m.subject, m.preview, m.folder, m.sender && m.sender.name, m.sender && m.sender.address,
      ...(m.to || []).map((r) => `${r.name} ${r.address}`)].join(" ");
    if (fold(haystack).includes(needle)) return true;
  }
  return false;
}

function render() {
  const needle = fold($("filter").value.trim());
  const container = $("conversations");
  container.replaceChildren(...sortedConversations().filter((t) => matchesFilter(t, needle)).map(renderConversation));
  if (!container.children.length) container.append(el("p", { class: "muted", style: "padding:0 12px" }, "No messages."));
  $("more").hidden = !state.cursor;
  renderSelection();
}

function isSingle(conversation) {
  if (conversation.messages.size !== 1) return false;
  if (!conversation.conversationId) return true;
  return conversation.size !== undefined && !conversation.sizeAtLeast && conversation.size <= 1;
}

function countLabel(conversation, loaded) {
  if (conversation.complete) return `${loaded}`;
  if (conversation.size !== undefined) return `${Math.max(conversation.size, loaded)}${conversation.sizeAtLeast ? "+" : ""}`;
  return `${loaded}+`;
}

// Rows look like Outlook's: sender, subject, preview; a blue bar and blue subject when unread; flag
// and paperclip icons; file chips; meeting mail labelled, with its time and place. The whole row is
// clickable; only a file chip has its own action (download).

const MEETING_LABELS = { invite: "Invite", update: "Updated", cancelled: "Canceled", accepted: "Accepted",
  tentative: "Tentative", declined: "Declined" };
const RESPONSES = new Set(["accepted", "tentative", "declined"]);

function rowClasses(base, { unread, active, response, matched }) {
  return [base, unread && "unread", active && "active", response && "response", matched && "matched"].filter(Boolean).join(" ");
}

function kindBadge(meeting) {
  return meeting ? el("span", { class: `kind kind-${meeting.kind}` }, MEETING_LABELS[meeting.kind]) : null;
}

function subjectLine(message) {
  const cancelled = message.meeting && message.meeting.kind === "cancelled";
  return el("div", { class: "subject" }, kindBadge(message.meeting),
    el("span", { class: cancelled ? "struck" : "" }, message.subject || "(no subject)"));
}

// "Tue Oct 6, 10:00 – 10:30 AM · Microsoft Teams Meeting"; a cancellation says when it was
function meetingPanel(meeting) {
  if (!meeting || RESPONSES.has(meeting.kind) || !meeting.start) return null;
  const start = new Date(meeting.start);
  const end = meeting.end ? new Date(meeting.end) : null;
  const day = start.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
  const hours = (d) => d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  const time = meeting.all_day ? `${day} (all day)` : `${day}, ${hours(start)}${end ? ` – ${hours(end)}` : ""}`;
  const cancelled = meeting.kind === "cancelled";
  const text = [cancelled ? `Was ${time}` : time, meeting.location].filter(Boolean).join(" · ");
  return el("div", { class: `meeting${cancelled ? " cancelled" : ""}${meeting.out_of_date ? " outdated" : ""}`,
    title: meeting.out_of_date ? "A newer update replaced this invitation" : undefined },
    icon(cancelled ? "calendarOff" : "calendar"), el("span", {}, text));
}

const FILE_COLORS = { xlsx: "sheet", xls: "sheet", csv: "sheet", docx: "doc", doc: "doc", pptx: "slides", ppt: "slides",
  pdf: "pdf", png: "image", jpg: "image", jpeg: "image", gif: "image", zip: "archive", eml: "mail", msg: "mail" };

function fileChips(message) {
  const files = state.attachments.get(message.id);
  if (!files || !files.length) return null;
  return el("div", { class: "chips" }, files.map((file) => {
    const name = file.name || "attachment";
    const kind = FILE_COLORS[name.split(".").pop().toLowerCase()] || "other";
    if (file.kind === "reference") return el("span", { class: "chip", title: "Cloud link" }, icon("file", kind), name);
    return el("button", { type: "button", class: "chip", title: `Download ${name}`,
      onclick: () => download(`/api/messages/${encodeURIComponent(message.id)}/attachments/${encodeURIComponent(file.id)}`, {}, "attachment").catch(() => {}) },
      icon("file", kind), name);
  }));
}

function sideColumn(date, { flagged, attachments }) {
  return el("div", { class: "side" },
    el("span", { class: "date" }, formatDate(date)),
    el("span", { class: "icons" }, flagged ? icon("flag", "flagged") : null, attachments ? icon("paperclip") : null));
}

// Ctrl+click (Cmd+click on a Mac) toggles a row's selection instead of opening it.
function clickRow(event, checkbox, open) {
  if (event.ctrlKey || event.metaKey) {
    if (!checkbox.disabled) checkbox.click();
    return;
  }
  open();
}

function rowAction(checkbox, open, attrs, ...children) {
  return el("button", { type: "button", class: "row-action", ...attrs,
    onclick: (event) => clickRow(event, checkbox, open) }, ...children);
}

function selectBox(title, checked, disabled, onchange) {
  const box = el("input", { type: "checkbox", title, "aria-label": title, disabled, onchange });
  box.checked = checked;
  return box;
}

// The invitation a conversation is about: the newest current one, else the newest cancellation.
function conversationMeeting(messages) {
  const meetings = messages.map((m) => m.meeting).filter(Boolean);
  return meetings.find((m) => (m.kind === "invite" || m.kind === "update") && !m.out_of_date)
    || meetings.find((m) => m.kind === "cancelled") || null;
}

// A conversation with one message: a plain row, like Outlook's conversation view.
function renderSingle(conversation) {
  const [message] = conversation.messages.values();
  const checkbox = selectBox("Export this message",
    state.selectedMessages.has(message.id) || state.selectedConversations.has(conversation.conversationId), false, (event) => {
      toggle(state.selectedMessages, message.id, event.target.checked);
      if (!event.target.checked && conversation.conversationId) state.selectedConversations.delete(conversation.conversationId);
      render();
    });
  const response = message.meeting && RESPONSES.has(message.meeting.kind);
  return el("div", { class: "conversation" },
    el("div", { class: rowClasses("conversation-row", { unread: message.is_read === false, active: state.activeMessage === message.id, response }) },
      checkbox,
      el("div", { class: "row-content" },
        rowAction(checkbox, () => openMessage(message.id), { class: "row-action conversation-action",
          "aria-label": `Open message: ${message.subject || "(no subject)"}` },
          el("span", { class: "toggle", "aria-hidden": "true" }),
          el("div", { class: "lines" },
            el("div", { class: "who" }, who(message.sender), ...badges(message, state.folder === null)),
            subjectLine(message),
            response ? null : el("div", { class: "preview" }, message.preview || ""),
            meetingPanel(message.meeting)),
          sideColumn(message.received_at, { flagged: message.flagged, attachments: message.has_attachments })),
        fileChips(message))));
}

function renderConversation(conversation) {
  if (isSingle(conversation)) return renderSingle(conversation);
  // newest on top, like the list and Outlook's conversation view (exports and get_conversation stay oldest first)
  const when = (m) => Date.parse(m.received_at || 0);
  const messages = [...conversation.messages.values()].sort((a, b) => when(b) - when(a));
  const newest = messages[0];
  const senders = [...new Set(messages.map((m) => who(m.sender)))].join(SEPARATOR);
  const unread = messages.some((m) => m.is_read === false);
  const selectable = Boolean(conversation.conversationId);
  const checkbox = selectBox("Export the whole conversation", state.selectedConversations.has(conversation.conversationId), !selectable,
    (event) => { toggle(state.selectedConversations, conversation.conversationId, event.target.checked); render(); });
  const meeting = conversationMeeting(messages);
  const row = el("div", { class: rowClasses("conversation-row", { unread }) },
    checkbox,
    rowAction(checkbox, () => expand(conversation), { class: "row-action conversation-action",
      "aria-label": `${conversation.expanded ? "Collapse" : "Expand"} conversation: ${newest.subject || "(no subject)"}`,
      "aria-expanded": String(conversation.expanded) },
      el("span", { class: "toggle", "aria-hidden": "true" }, icon(conversation.expanded ? "chevronDown" : "chevronRight")),
      el("div", { class: "lines" },
        el("div", { class: "who" }, senders, el("span", { class: "conversation-count" }, countLabel(conversation, messages.length))),
        el("div", { class: "subject" }, kindBadge(meeting), el("span", {}, newest.subject || "(no subject)")),
        el("div", { class: "preview" }, newest.preview || ""),
        meetingPanel(meeting)),
      sideColumn(newest.received_at,
        { flagged: messages.some((m) => m.flagged), attachments: messages.some((m) => m.has_attachments) })));
  const node = el("div", { class: "conversation" }, row);
  if (conversation.expanded) {
    node.append(el("div", { class: "messages" }, messages.map((m) => renderMessage(m, conversation)),
      conversation.loading ? spinner("Loading the whole conversation…") : null));
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

function renderMessage(message, conversation) {
  const covered = state.selectedConversations.has(conversation.conversationId);
  const checkbox = selectBox(covered ? "Included with the conversation" : "Export this message",
    covered || state.selectedMessages.has(message.id), covered,
    (event) => { toggle(state.selectedMessages, message.id, event.target.checked); renderSelection(); });
  const response = message.meeting && RESPONSES.has(message.meeting.kind);
  return el("div", { class: rowClasses("msg-row", { unread: message.is_read === false, active: state.activeMessage === message.id, response, matched: message.matched }),
    title: message.matched ? "Matches the search" : undefined },
    checkbox,
    el("div", { class: "row-content" },
      rowAction(checkbox, () => openMessage(message.id), { class: "row-action message-action",
        "aria-label": `Open message: ${message.subject || "(no subject)"}` },
        el("div", { class: "lines" },
          el("div", { class: "who" }, kindBadge(message.meeting), who(message.sender), ...badges(message, true)),
          response ? null : el("div", { class: "preview" }, message.preview || message.subject || ""),
          meetingPanel(message.meeting)),
        sideColumn(message.received_at, { flagged: message.flagged, attachments: message.has_attachments })),
      fileChips(message)));
}

function toggle(set, value, on) {
  if (on) set.add(value); else set.delete(value);
}

async function expand(conversation) {
  conversation.expanded = !conversation.expanded;
  if (conversation.expanded && conversation.conversationId && !conversation.complete && !conversation.loading) {
    clearBanner("conversation");
    conversation.loading = true;
    render();
    try {
      const full = await json(`/api/conversations/${encodeURIComponent(conversation.conversationId)}?${query({ deleted_items: showDeletedItems() })}`,
        undefined, "conversation", () => state.conversations.get(conversation.key) === conversation);
      for (const entry of full.messages) conversation.messages.set(entry.message.id, { ...entry.message, matched: conversation.messages.get(entry.message.id)?.matched });
      // incomplete coverage: the conversation is larger than the server lists (a "1000+" conversation)
      conversation.complete = full.coverage.complete;
      conversation.size = full.messages.length;
      conversation.sizeAtLeast = !full.coverage.complete;
      loadAttachmentNames();
    } catch {
      // The API helper shows the error; collapsing and reopening the row can retry it.
    } finally {
      conversation.loading = false;
    }
  }
  if (state.conversations.get(conversation.key) === conversation) render(); // skip if the list was replaced meanwhile
}

// ------------------------------------------------------------------ reader

async function openMessage(id) {
  clearBanner("reader");
  state.activeMessage = id;
  render();
  const request = ++state.readerRequest;
  $("reader-title").replaceChildren(spinner("Loading message…"));
  $("reader-meta").replaceChildren();
  $("reader-body").textContent = "";
  const body = $("opt-full").checked ? "full" : "unique"; // one setting for the reader and exports
  let content;
  try {
    content = await json(`/api/messages/${encodeURIComponent(id)}?${query({ body })}`, undefined, "reader",
      () => request === state.readerRequest); // the whole body
  } catch {
    if (request === state.readerRequest) $("reader-title").textContent = "Could not load the complete message.";
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
    onclick: () => download(`/api/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachment.id)}`, {}, "attachment").catch(() => {}) }, label);
}

// Fetch with the session token and hand the response to the browser as a download.
async function download(path, options, owner = null) {
  clearBanner(owner);
  const response = await api(path, options, owner);
  const url = URL.createObjectURL(await response.blob());
  const link = el("a", { href: url, download: filenameFrom(response.headers.get("Content-Disposition")) });
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
  return response;
}

// ------------------------------------------------------------------ selection and export

function renderSelection() {
  const conversations = state.selectedConversations.size;
  const messages = [...state.selectedMessages].filter((id) => !coveredByConversation(id)).length;
  const parts = [];
  if (conversations) parts.push(`${conversations} conversation${conversations > 1 ? "s" : ""}`);
  if (messages) parts.push(`${messages} message${messages > 1 ? "s" : ""}`);
  $("selection").textContent = parts.length ? `Selected: ${parts.join(" + ")}` : "Nothing selected";
  $("export").disabled = !parts.length;
}

function coveredByConversation(messageId) {
  for (const conversation of state.conversations.values()) {
    if (conversation.messages.has(messageId) && state.selectedConversations.has(conversation.conversationId)) return true;
  }
  return false;
}

function exportOptions() {
  return {
    include_attachments: $("opt-attachments").checked,
    combine: document.querySelector('input[name="files"]:checked').value,
    body: $("opt-full").checked ? "full" : "unique",
    scope: { deleted_items: showDeletedItems() },
  };
}

function exportRequest() {
  return {
    conversation_ids: [...state.selectedConversations],
    message_ids: [...state.selectedMessages].filter((id) => !coveredByConversation(id)),
    ...exportOptions(),
  };
}

// The whole current view: this folder (or the mailbox) within the chosen dates, up to 2,000 messages.
function viewRequest() {
  const { since, until } = dateBounds();
  const options = exportOptions();
  return { ...options, folder: state.folder, since, until,
    scope: { ...options.scope, meeting_mail: $("opt-meetings").checked } };
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
    const response = await download("/api/export", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(request) }, "export");
    // Parts that could not be exported: the file's header line, shown until the next export retry.
    const errors = response.headers.get("X-Export-Errors");
    if (errors) showBanner(decodeURIComponent(errors).replace("below", "in the file"), "export");
  } catch { /* the banner shows the error */ } finally {
    button.textContent = label;
    renderSelection();
    renderExportView();
  }
}

// ------------------------------------------------------------------ wiring

function clearMailSearch() {
  const wasSearching = state.mode === "search";
  state.query = "";
  if (!wasSearching) return;
  state.mode = "list";
  $("list-title").textContent = state.folderName;
  renderExportView();
  loadList(true);
}

bindClearableSearch($("search"), $("search-clear"), { onClear: clearMailSearch });
bindExpandableSearch({
  host: $("sidebar-head"),
  toggle: $("folder-search-toggle"),
  input: $("folder-filter"),
  clearButton: $("folder-filter-clear"),
  onInput: renderFolderFilter,
  hideTitle: true,
});
bindExpandableSearch({
  host: $("conversation-search-head"),
  toggle: $("conversation-search-toggle"),
  input: $("filter"),
  clearButton: $("filter-clear"),
  onInput: render,
});

$("search-form").addEventListener("submit", (event) => {
  event.preventDefault();
  state.query = $("search").value.trim();
  state.mode = state.query ? "search" : "list";
  $("list-title").textContent = state.query ? `Search: ${state.query}` : state.folderName;
  renderExportView();
  (state.query ? runSearch : loadList)(true);
});

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

function focusCalendarDate(key) {
  $("calendar-days").querySelector(`[data-date="${key}"]`)?.focus();
}

function setCalendarOpen(open, { restoreFocus = false } = {}) {
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
    focusCalendarDate($("since").value || dateKey(new Date()));
    return;
  }
  const range = `${$("since").value}|${$("until").value}`;
  if (range !== picker.applied) { // reload only when the range actually changed
    picker.applied = range;
    renderExportView();
    (state.mode === "search" ? runSearch : loadList)(true);
  }
  if (restoreFocus) $("range-toggle").focus();
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
  const today = dateKey(new Date());
  const days = [];
  for (let index = 0; index < 42; index += 1) {
    const day = new Date(first.getFullYear(), first.getMonth(), 1 - mondayOffset + index);
    const key = dateKey(day);
    const future = key > today;
    const classes = ["calendar-day"];
    if (day.getMonth() !== picker.month.getMonth()) classes.push("outside-month");
    if (start && end && key >= start && key <= end) classes.push("in-range");
    if (start && key === start) classes.push("range-start");
    if (end && key === end) classes.push("range-end");
    days.push(el("button", { type: "button", class: classes.join(" "), "data-date": key,
      "aria-label": day.toLocaleDateString(undefined, { dateStyle: "full" }),
      "aria-pressed": String(Boolean(start && end && key >= start && key <= end)),
      disabled: future,
      onclick: () => chooseDate(key) }, day.getDate()));
  }
  $("calendar-days").replaceChildren(...days);
}

function chooseDate(key) {
  if (key > dateKey(new Date())) return;
  if (!picker.choosingEnd || key < picker.draftStart) {
    picker.draftStart = key;
    picker.choosingEnd = true;
    setDateRange(key, key);
    focusCalendarDate(key);
    return;
  }
  setDateRange(picker.draftStart, key);
  picker.choosingEnd = false;
  setCalendarOpen(false, { restoreFocus: true });
}

$("range-toggle").addEventListener("click", () => setCalendarOpen($("calendar").hidden));
$("previous-month").addEventListener("click", () => { picker.month.setMonth(picker.month.getMonth() - 1); renderCalendar(); });
$("next-month").addEventListener("click", () => { picker.month.setMonth(picker.month.getMonth() + 1); renderCalendar(); });
$("clear-dates").addEventListener("click", () => { picker.choosingEnd = false; setDateRange("", ""); setCalendarOpen(false, { restoreFocus: true }); });
$("calendar-done").addEventListener("click", () => setCalendarOpen(false, { restoreFocus: true }));
document.addEventListener("pointerdown", (event) => {
  if (!$("calendar").hidden && !$("range-picker").contains(event.target)) setCalendarOpen(false);
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !$("calendar").hidden) setCalendarOpen(false, { restoreFocus: true });
});
$("more").addEventListener("click", () => (state.mode === "search" ? runSearch : loadList)(false));
$("export-view").addEventListener("click", () => runExport($("export-view"), viewRequest(), "Export view"));
$("refresh-folders").addEventListener("click", () => loadFolders(true));
$("opt-full").addEventListener("change", () => state.activeMessage && openMessage(state.activeMessage));
$("opt-deleted").addEventListener("change", () => (state.mode === "search" ? runSearch : loadList)(true));
$("opt-meetings").addEventListener("change", () => (state.mode === "search" ? runSearch : loadList)(true));
$("clear").addEventListener("click", () => { state.selectedConversations.clear(); state.selectedMessages.clear(); render(); });
$("export").addEventListener("click", () => runExport($("export"), exportRequest(), "Export"));
setInterval(() => {
  clearBanner("heartbeat");
  api("/api/heartbeat", { method: "POST" }, "heartbeat").catch(() => {});
}, 60_000);

(async function start() {
  let status;
  clearBanner("status");
  try {
    status = await json("/api/status", undefined, "status");
  } catch {
    return; // the API helper has shown the startup error
  }
  $("account-name").textContent = status.account || "";
  if (!status.signed_in.graph) {
    showBanner(`Not signed in. Run \`${status.sign_in_command}\` in a terminal, then reload this page.`);
    return;
  }
  renderExportView();
  loadProfile(); // the header's name and photo; never blocks the mail
  await Promise.all([loadFolders(), loadList(true)]); // independent: load side by side
})();

// Your display name and photo in the header. Plain fetches: a failure here only keeps the address
// and the initials, and never touches the error banner.
async function loadProfile() {
  const headers = { "X-Session-Token": TOKEN };
  try {
    const response = await fetch("/api/me", { headers });
    if (!response.ok) return;
    const me = await response.json();
    const name = me.display_name || me.email || "";
    $("account-name").textContent = name;
    $("account").title = me.email || "";
    $("avatar").textContent = initials(name);
    const photo = await fetch("/api/me/photo", { headers });
    if (photo.ok) $("avatar").replaceChildren(el("img", { src: URL.createObjectURL(await photo.blob()), alt: "" }));
  } catch { /* the address stays */ }
}

// "Doe, Jane" (Outlook's "Last, First") or "Jane Doe" -> "JD"
function initials(name) {
  const parts = name.includes(",") ? name.split(",").reverse() : name.split(/\s+/);
  return parts.map((part) => part.trim()[0] || "").join("").slice(0, 2).toUpperCase();
}
