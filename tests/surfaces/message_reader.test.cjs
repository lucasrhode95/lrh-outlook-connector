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
  setAttribute() {}
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
      if (path === '/api/status') return { ok: true, json: async () => ({ signed_in: { read: false } }) };
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
