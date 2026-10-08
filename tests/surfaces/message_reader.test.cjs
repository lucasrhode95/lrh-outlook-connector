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
    this.listeners = new Map();
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
  addEventListener(type, listener) {
    const listeners = this.listeners.get(type) || [];
    listeners.push(listener);
    this.listeners.set(type, listeners);
  }
  dispatch(type, event) {
    for (const listener of this.listeners.get(type) || []) listener(event);
  }
  focus() {}
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
  const timeouts = [];
  const context = vm.createContext({
    document: {
      querySelector: () => ({ content: 'synthetic-token' }),
      getElementById: node,
      createElement: () => new Element(),
      createTextNode: (text) => text,
      addEventListener() {},
    },
    Node: Element, URLSearchParams, setInterval() {},
    setTimeout: callback => { timeouts.push(callback); return timeouts.length; },
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
  return { context, nodes, calls, timeouts, open: (id) => context.openMessage(id) };
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
  assert.equal(vm.runInContext('picker.choosingEnd', ui.context), false);
});

test('calendar grid starts on Sunday', () => {
  const ui = reader(async () => ({}));
  ui.context.monthYear = 2025;
  ui.context.monthIndex = 8; // September 2025 starts on Monday.
  vm.runInContext('picker.month = new Date(monthYear, monthIndex, 1); renderCalendar();', ui.context);
  const dates = ui.nodes.get('calendar-days').children.slice(0, 2).map(day => day.attributes.get('data-date'));
  assert.deepEqual(dates, ['2025-08-31', '2025-09-01']);
});

test('selecting the end date before the start reorders the range', () => {
  const ui = reader(async () => ({}));
  const today = new Date();
  const key = date => String(date.getFullYear()).padStart(4, '0') + '-' +
    String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
  const later = key(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1));
  const earlier = key(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 4));
  ui.context.laterDate = later;
  ui.context.earlierDate = earlier;
  vm.runInContext('chooseDate(laterDate); chooseDate(earlierDate);', ui.context);
  assert.equal(ui.nodes.get('since').value, earlier);
  assert.equal(ui.nodes.get('until').value, later);
  assert.equal(vm.runInContext('picker.choosingEnd', ui.context), false);
});

test('Today navigates to the current month without losing a pending date selection', () => {
  const ui = reader(async () => ({}));
  const today = new Date();
  const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
  const start = String(yesterday.getFullYear()).padStart(4, '0') + '-' +
    String(yesterday.getMonth() + 1).padStart(2, '0') + '-' + String(yesterday.getDate()).padStart(2, '0');
  ui.context.selectedDate = start;
  vm.runInContext('chooseDate(selectedDate); picker.month = new Date(2020, 0, 1); renderCalendar(); $("calendar").hidden = false;', ui.context);
  ui.nodes.get('today-date').dispatch('click', {});
  assert.equal(ui.nodes.get('calendar-month').textContent,
    today.toLocaleDateString(undefined, { month: 'long', year: 'numeric' }));
  assert.equal(ui.nodes.get('since').value, start);
  assert.equal(ui.nodes.get('until').value, start);
  assert.equal(vm.runInContext('picker.draftStart', ui.context), start);
  assert.equal(vm.runInContext('picker.choosingEnd', ui.context), true);
  assert.equal(ui.nodes.get('calendar').hidden, false);
});

test('date filter clear button clears the range and closes the calendar', () => {
  const ui = reader(async () => ({}));
  ui.context.selectedDate = '2025-03-12';
  vm.runInContext('chooseDate(selectedDate); picker.applied = "|";', ui.context);
  const clearButton = ui.nodes.get('date-filter-clear');
  assert.equal(clearButton.hidden, false);
  ui.nodes.get('calendar').hidden = false;
  clearButton.dispatch('click', {});
  assert.equal(ui.nodes.get('since').value, '');
  assert.equal(ui.nodes.get('until').value, '');
  assert.equal(ui.nodes.get('range-label').textContent, 'Any date');
  assert.equal(clearButton.hidden, true);
  assert.equal(ui.nodes.get('calendar').hidden, true);
  assert.equal(vm.runInContext('picker.choosingEnd', ui.context), false);
});

test('wheel and horizontal swipes navigate calendar months in the expected direction', () => {
  const ui = reader(async () => ({}));
  const today = new Date();
  const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const monthName = date => date.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
  ui.context.monthYear = firstOfMonth.getFullYear();
  ui.context.monthIndex = firstOfMonth.getMonth();
  vm.runInContext('picker.month = new Date(monthYear, monthIndex, 1); renderCalendar();', ui.context);
  const calendar = ui.nodes.get('calendar');
  calendar.hidden = false;
  const monthLabel = () => ui.nodes.get('calendar-month').textContent;
  let prevented = false;

  calendar.dispatch('wheel', { deltaY: -80, preventDefault() { prevented = true; } });
  const previousMonth = new Date(firstOfMonth.getFullYear(), firstOfMonth.getMonth() - 1, 1);
  assert.equal(monthLabel(), monthName(previousMonth));
  assert.equal(prevented, true);
  ui.timeouts.shift()();
  calendar.dispatch('wheel', { deltaY: 80, preventDefault() {} });
  assert.equal(monthLabel(), monthName(firstOfMonth));
  ui.timeouts.shift()();

  calendar.dispatch('touchstart', { touches: [{ clientX: 100, clientY: 50 }] });
  calendar.dispatch('touchend', { changedTouches: [{ clientX: 160, clientY: 55 }] });
  assert.equal(monthLabel(), monthName(previousMonth));
  calendar.dispatch('touchstart', { touches: [{ clientX: 160, clientY: 55 }] });
  calendar.dispatch('touchend', { changedTouches: [{ clientX: 100, clientY: 50 }] });
  assert.equal(monthLabel(), monthName(firstOfMonth));
});
