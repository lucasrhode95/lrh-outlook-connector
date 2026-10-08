// Behavioral reader tests using the shipped app, Node's built-in runner and a minimal DOM.
const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const vm = require('node:vm');

const source = readFileSync(join(__dirname, '../../src/outlook_connector/surfaces/web/static/app.js'), 'utf8');

class Element {
  constructor() {
    this.textContent = ''; this.children = []; this.checked = false; this.value = '';
    this.attributes = new Map(); this.disabled = false;
    const classes = new Set();
    this.classList = {
      add: (...names) => names.forEach(name => classes.add(name)),
      remove: (...names) => names.forEach(name => classes.delete(name)),
      contains: name => classes.has(name),
      toggle: (name, force) => {
        const shouldAdd = force ?? !classes.has(name);
        if (shouldAdd) classes.add(name);
        else classes.delete(name);
        return shouldAdd;
      },
    };
  }
  addEventListener() {}
  querySelector() { return { focus() {} }; }
  setAttribute(name, value) {
    this.attributes.set(name, value);
    if (name === 'disabled') this.disabled = true;
  }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
}

function reader(handler, full = false) {
  const nodes = new Map();
  const node = (id) => {
    if (!nodes.has(id)) nodes.set(id, new Element());
    return nodes.get(id);
  };
  node('opt-full').checked = full;
  const calls = [];
  const context = vm.createContext({
    document: {
      querySelector: () => ({ content: 'synthetic-token' }),
      getElementById: node,
      createElement: () => new Element(),
      createTextNode: (text) => text,
      addEventListener() {},
    },
    Node: Element, URLSearchParams, setInterval() {},
    fetch: async (path, options) => {
      if (path === '/api/status') return { ok: true, json: async () => ({ signed_in: { graph: false } }) };
      assert.equal(options.headers['X-Session-Token'], 'synthetic-token');
      const url = new URL(path, 'http://localhost');
      calls.push(url);
      const data = await handler(url);
      return { ok: true, json: async () => data };
    },
  });
  vm.runInContext(source, context);
  // Isolate reader behavior from list rendering; no production test seams.
  vm.runInContext('render = () => {};', context);
  return { context, nodes, calls, open: (id) => context.openMessage(id) };
}

function part(text, id = 'mail') {
  return { text, offset: 0, message: { id, subject: id }, attachments: [] };
}

test('loads the whole body in one request with the chosen body kind', async () => {
  const ui = reader(async () => part('Hello😀\n  middle \nEnd'), true);
  await ui.open('mail');
  assert.equal(ui.nodes.get('reader-body').textContent, 'Hello😀\n  middle \nEnd');
  assert.equal(ui.calls.length, 1);
  assert.equal(ui.calls[0].searchParams.get('body'), 'full');
  assert.equal(ui.calls[0].searchParams.get('offset'), null);
});

test('has no body-size cap: a very long unique body is shown whole from one request', async () => {
  const text = 'x'.repeat(200000) + 'y'.repeat(200000) + 'z'.repeat(200001);
  const ui = reader(async () => part(text, 'large'));
  await ui.open('large');
  assert.equal(ui.nodes.get('reader-body').textContent, text);
  assert.equal(ui.calls.length, 1);
  assert.equal(ui.calls[0].searchParams.get('body'), 'unique');
});

test('a failure never presents a partial body', async () => {
  const ui = reader(async () => { throw new Error('synthetic failure'); });
  await ui.open('broken');
  assert.equal(ui.nodes.get('reader-title').textContent, 'Could not load the complete message.');
  assert.equal(ui.nodes.get('reader-body').textContent, '');
});

test('switching messages keeps the newest body when an older answer arrives late', async () => {
  let resolveOld;
  const old = new Promise(resolve => { resolveOld = resolve; });
  const ui = reader(async (url) => url.pathname.endsWith('/old') ? old : part('Newest', 'new'));
  const readingOld = ui.open('old');
  await ui.open('new');
  resolveOld(part('Old body', 'old'));
  await readingOld;
  assert.equal(ui.nodes.get('reader-body').textContent, 'Newest');
  assert.equal(ui.calls.filter(u => u.pathname.endsWith('/old')).length, 1);
});

test('future calendar dates are disabled and cannot be selected', () => {
  const ui = reader(async () => ({}));
  const today = new Date();
  const key = date => String(date.getFullYear()).padStart(4, '0') + '-' +
    String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
  const todayKey = key(today);
  ui.context.monthYear = today.getFullYear();
  ui.context.monthIndex = today.getMonth();
  vm.runInContext('picker.month = new Date(monthYear, monthIndex, 1); renderCalendar();', ui.context);
  const todayButton = ui.nodes.get('calendar-days').children.find(day => day.attributes.get('data-date') === todayKey);
  assert.ok(todayButton);
  assert.equal(todayButton.disabled, false);

  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
  const tomorrowKey = key(tomorrow);
  ui.context.monthYear = tomorrow.getFullYear();
  ui.context.monthIndex = tomorrow.getMonth();
  vm.runInContext('picker.month = new Date(monthYear, monthIndex, 1); renderCalendar();', ui.context);
  const tomorrowButton = ui.nodes.get('calendar-days').children.find(day => day.attributes.get('data-date') === tomorrowKey);
  assert.ok(tomorrowButton);
  assert.equal(tomorrowButton.disabled, true);
  ui.context.selectedDate = tomorrowKey;
  vm.runInContext('chooseDate(selectedDate);', ui.context);
  assert.equal(ui.nodes.get('since').value, '');
  assert.equal(ui.nodes.get('until').value, '');
});

test('selecting an end date leaves the date picker open', () => {
  const ui = reader(async () => ({}));
  const today = new Date();
  const key = date => String(date.getFullYear()).padStart(4, '0') + '-' +
    String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
  const start = key(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 2));
  const end = key(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1));
  vm.runInContext('$("calendar").hidden = false;', ui.context);
  ui.context.selectedDate = start;
  vm.runInContext('chooseDate(selectedDate);', ui.context);
  ui.context.selectedDate = end;
  vm.runInContext('chooseDate(selectedDate);', ui.context);
  assert.equal(ui.nodes.get('since').value, start);
  assert.equal(ui.nodes.get('until').value, end);
  assert.equal(ui.nodes.get('calendar').hidden, false);
  assert.match(ui.nodes.get('calendar-hint').textContent, /Range selected/);
});
