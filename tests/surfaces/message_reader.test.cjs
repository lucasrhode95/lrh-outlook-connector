// Behavioral reader tests using the shipped app, Node's built-in runner and a minimal DOM.
const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const vm = require('node:vm');

const source = readFileSync(join(__dirname, '../../src/outlook_connector/surfaces/web/static/app.js'), 'utf8');

class Element {
  constructor() { this.textContent = ''; this.children = []; this.checked = false; this.value = ''; }
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

function part(text, offset, next_offset, id = 'mail') {
  return { text, offset, next_offset, message: { id, subject: id }, attachments: [] };
}

test('automatically loads all continuations with exact server offsets and the chosen body', async () => {
  const chunks = [part('Hello😀', 0, 6), part('\n  middle\u00a0', 6, 17), part('\nEnd', 17, null)];
  const ui = reader(async (url) => chunks[Number(url.searchParams.get('offset')) === 0 ? 0 :
    Number(url.searchParams.get('offset')) === 6 ? 1 : 2], true);
  await ui.open('mail');
  assert.equal(ui.nodes.get('reader-body').textContent, 'Hello😀\n  middle\u00a0\nEnd');
  assert.deepEqual(ui.calls.map(u => u.searchParams.get('offset')), ['0', '6', '17']);
  assert.ok(ui.calls.every(u => u.searchParams.get('body') === 'full'));
});

test('has no body-size cap and loads unique bodies completely', async () => {
  const texts = ['x'.repeat(200000), 'y'.repeat(200000), 'z'.repeat(200001)];
  let i = 0;
  const ui = reader(async () => { const index = i++; return part(texts[index], index * 200000,
    index < 2 ? (index + 1) * 200000 : undefined); });
  await ui.open('large');
  assert.equal(ui.nodes.get('reader-body').textContent, texts.join(''));
  assert.equal(ui.calls.length, 3);
  assert.ok(ui.calls.every(u => u.searchParams.get('body') === 'unique'));
});

test('a complete short body needs only one request', async () => {
  const ui = reader(async () => part('Short', 0, undefined));
  await ui.open('short');
  assert.equal(ui.nodes.get('reader-body').textContent, 'Short');
  assert.equal(ui.calls.length, 1);
});

test('continuation failure never presents a partial body as complete', async () => {
  const ui = reader(async (url) => {
    if (url.searchParams.get('offset') === '0') return part('Partial', 0, 7);
    throw new Error('synthetic failure');
  });
  await ui.open('broken');
  assert.equal(ui.nodes.get('reader-title').textContent, 'Could not load the complete message.');
  assert.equal(ui.nodes.get('reader-body').textContent, '');
});

test('switching messages stops obsolete continuations and keeps the newest body', async () => {
  let resolveOld;
  const old = new Promise(resolve => { resolveOld = resolve; });
  const ui = reader(async (url) => url.pathname.endsWith('/old') ? old : part('Newest', 0, null, 'new'));
  const readingOld = ui.open('old');
  await ui.open('new');
  resolveOld(part('Old partial', 0, 11, 'old'));
  await readingOld;
  assert.equal(ui.nodes.get('reader-body').textContent, 'Newest');
  assert.equal(ui.calls.filter(u => u.pathname.endsWith('/old')).length, 1);
});
