// Outlook connector local UI. Mail content is untrusted: it is only ever inserted as text.

const TOKEN = document.querySelector('meta[name="session-token"]').content;
const $ = (id) => document.getElementById(id);

const state = {
  folder: null,          // folder id or null for the whole mailbox
  mode: "list",          // "list" | "search"
  query: "",
  cursor: null,
  threads: new Map(),    // key -> { key, conversationId, subject, messages: Map(id -> summary), expanded, complete }
  selectedThreads: new Set(),
  selectedMessages: new Set(),
  activeMessage: null,
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
  const folders = await json(`/api/folders?${query({ refresh })}`);
  const list = $("folders");
  list.replaceChildren(folderItem(null, "All mail (recent)", 0, null));
  for (const folder of folders) {
    if (folder.hidden) continue;
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
    state.mode = "list";
    $("list-title").textContent = id ? name : "Recent mail";
    for (const li of $("folders").children) li.classList.toggle("active", li === item);
    loadList(true);
  });
  return item;
}

// ------------------------------------------------------------------ list and search

function resetThreads() {
  state.threads = new Map();
  state.cursor = null;
}

function addMessage(summary, { matched = false } = {}) {
  const key = summary.conversation_id || summary.id;
  let thread = state.threads.get(key);
  if (!thread) {
    thread = { key, conversationId: summary.conversation_id, subject: summary.subject, messages: new Map(), expanded: false, complete: false };
    state.threads.set(key, thread);
  }
  thread.messages.set(summary.id, { ...summary, matched });
}

async function loadList(reset) {
  if (reset) resetThreads();
  $("coverage").textContent = "Loading…";
  const { since, until } = dateBounds();
  const page = await json(`/api/messages?${query({ folder: state.folder, since, until, limit: 100, cursor: state.cursor })}`);
  for (const item of page.items) addMessage(item);
  state.cursor = page.cursor;
  showCoverage(page.coverage);
  render();
}

async function runSearch(reset) {
  if (reset) resetThreads();
  $("coverage").textContent = "Searching…";
  const { since, until } = dateBounds();
  const result = await json(`/api/search?${query({ q: state.query, since, until, folder: state.folder, limit: 50, cursor: state.cursor })}`);
  for (const hit of result.conversations) for (const message of hit.matching_messages) addMessage(message, { matched: true });
  state.cursor = result.cursor;
  showCoverage(result.coverage);
  render();
}

function showCoverage(coverage) {
  const parts = [coverage.complete ? "complete" : "more available"];
  if (coverage.server_total !== null && coverage.server_total !== undefined) parts.push(`${coverage.server_total} matching messages on the server`);
  if (coverage.source !== "remote") parts.push(coverage.source);
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

function renderThread(thread) {
  const messages = [...thread.messages.values()].sort((a, b) => Date.parse(a.received_at || 0) - Date.parse(b.received_at || 0));
  const newest = messages[messages.length - 1];
  const senders = [...new Set(messages.map((m) => who(m.sender)))].join(", ");
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
        el("span", { class: "badge" }, thread.complete ? `${messages.length}` : `${messages.length}+`),
        messages.some((m) => m.is_deleted) ? el("span", { class: "badge deleted" }, "deleted on server") : null),
      el("div", { class: "who" }, senders)),
    el("span", { class: "date" }, formatDate(newest.received_at || newest.sent_at)));
  const node = el("div", { class: "thread" }, row);
  if (thread.expanded) node.append(el("div", { class: "messages" }, messages.map((m) => renderMessage(m, thread))));
  return node;
}

function renderMessage(message, thread) {
  const covered = state.selectedThreads.has(thread.conversationId);
  const checkbox = el("input", { type: "checkbox", disabled: covered, title: covered ? "Included with the thread" : "Export this message",
    onclick: (event) => event.stopPropagation(),
    onchange: (event) => { toggle(state.selectedMessages, message.id, event.target.checked); renderSelection(); } });
  checkbox.checked = covered || state.selectedMessages.has(message.id);
  return el("div", { class: `msg-row${state.activeMessage === message.id ? " active" : ""}${message.is_read === false ? " unread" : ""}`,
    onclick: () => openMessage(message.id) },
    checkbox,
    el("div", {},
      el("div", { class: "msg-subject" }, who(message.sender),
        message.folder ? el("span", { class: "badge" }, message.folder) : null,
        message.is_deleted ? el("span", { class: "badge deleted" }, "deleted on server") : null),
      el("div", { class: "who" }, message.preview || message.subject || "")),
    el("span", { class: "date" }, formatDate(message.received_at || message.sent_at)));
}

function toggle(set, value, on) {
  if (on) set.add(value); else set.delete(value);
}

async function expand(thread) {
  thread.expanded = !thread.expanded;
  if (thread.expanded && thread.conversationId && !thread.complete) {
    const full = await json(`/api/threads/${encodeURIComponent(thread.conversationId)}?${query({ include_deleted_items: $("opt-deleted").checked })}`);
    for (const entry of full.messages) thread.messages.set(entry.message.id, { ...entry.message, matched: thread.messages.get(entry.message.id)?.matched });
    thread.complete = true;
  }
  render();
}

// ------------------------------------------------------------------ reader

async function openMessage(id) {
  state.activeMessage = id;
  render();
  $("reader-title").textContent = "Loading…";
  $("reader-meta").replaceChildren();
  $("reader-body").textContent = "";
  const body = $("reader-full").checked ? "full" : "unique";
  const content = await json(`/api/messages/${encodeURIComponent(id)}?${query({ body })}`);
  const m = content.message;
  $("reader-title").textContent = m.subject || "(no subject)";
  const rows = [["From", who(m.sender)], ["To", (m.to || []).map(who).join(", ")], ["Cc", (m.cc || []).map(who).join(", ")],
    ["Date", m.received_at ? new Date(m.received_at).toLocaleString() : ""], ["Folder", m.folder || ""],
    ["Attachments", (content.attachments || []).filter((a) => !a.is_inline).map((a) => a.name).join(", ")]];
  if (m.is_deleted) rows.push(["Note", "Deleted on the server; shown from local retention."]);
  $("reader-meta").replaceChildren(...rows.filter(([, v]) => v).flatMap(([k, v]) => [el("dt", {}, k), el("dd", {}, v)]));
  $("reader-body").textContent = content.text;
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

function exportRequest() {
  return {
    conversation_ids: [...state.selectedThreads],
    message_ids: [...state.selectedMessages].filter((id) => !coveredByThread(id)),
    include_attachments: $("opt-attachments").checked,
    combine: $("opt-all").checked ? "all" : ($("opt-per-thread").checked ? "per_thread" : "none"),
    body: $("opt-full").checked ? "full" : "unique",
    include_deleted_items: $("opt-deleted").checked,
  };
}

function filenameFrom(disposition) {
  const star = /filename\*=utf-8''([^;]+)/i.exec(disposition || "");
  if (star) return decodeURIComponent(star[1]);
  const plain = /filename="?([^";]+)"?/i.exec(disposition || "");
  return plain ? plain[1] : "outlook-export";
}

async function runExport() {
  const button = $("export");
  button.disabled = true;
  button.textContent = "Exporting…";
  try {
    const response = await api("/api/export", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(exportRequest()) });
    const url = URL.createObjectURL(await response.blob());
    const link = el("a", { href: url, download: filenameFrom(response.headers.get("Content-Disposition")) });
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60_000);
  } finally {
    button.textContent = "Export";
    renderSelection();
  }
}

// ------------------------------------------------------------------ wiring

$("search-form").addEventListener("submit", (event) => {
  event.preventDefault();
  state.query = $("search").value.trim();
  state.mode = state.query ? "search" : "list";
  $("list-title").textContent = state.query ? `Search: ${state.query}` : "Recent mail";
  (state.query ? runSearch : loadList)(true);
});
$("filter").addEventListener("input", render);
$("since").addEventListener("change", () => (state.mode === "search" ? runSearch : loadList)(true));
$("until").addEventListener("change", () => (state.mode === "search" ? runSearch : loadList)(true));
$("more").addEventListener("click", () => (state.mode === "search" ? runSearch : loadList)(false));
$("refresh-folders").addEventListener("click", () => loadFolders(true));
$("reader-full").addEventListener("change", () => state.activeMessage && openMessage(state.activeMessage));
$("opt-all").addEventListener("change", (event) => { $("opt-per-thread").disabled = event.target.checked; });
$("clear").addEventListener("click", () => { state.selectedThreads.clear(); state.selectedMessages.clear(); render(); });
$("export").addEventListener("click", runExport);
setInterval(() => api("/api/heartbeat", { method: "POST" }).catch(() => {}), 60_000);

(async function start() {
  const status = await json("/api/status");
  $("account").textContent = status.account || "";
  if (!status.signed_in.read) {
    showBanner(`Not signed in. Run \`${status.sign_in_command}\` in a terminal, then reload this page.`);
    return;
  }
  await loadFolders();
  await loadList(true);
})();
