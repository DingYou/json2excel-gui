const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'src/main.js'), 'utf8');

class Element {
  constructor(tag = 'div') { this.tagName = tag; this.children = []; this.style = {}; this.value = ''; this.listeners = {}; this.disabled = false; }
  appendChild(child) { this.children.push(child); return child; }
  replaceChildren(...children) { this.children = children; }
  addEventListener(event, handler) { this.listeners[event] = handler; }
  setAttribute(key, value) { this[key] = value; }
  querySelectorAll(selector) {
    const all = this.children.flatMap(child => [child, ...child.querySelectorAll(selector)]);
    if (selector.includes("input[type='checkbox']")) return all.filter(el => el.type === 'checkbox' && (!selector.endsWith(':checked') || el.checked));
    return [];
  }
  querySelector(selector) { return this.children.find(el => `.${el.className}` === selector || el.tagName === selector); }
  set innerHTML(_) { throw new Error('HTML injection sinks are not permitted'); }
}

function setup() {
  const elements = Object.fromEntries(['input-path', 'output-path', 'btn-convert', 'status-msg', 'btn-reveal', 'curl-input', 'json-input', 'field-selector', 'field-list'].map(id => [`#${id}`, new Element()]));
  const convert = elements['#btn-convert'];
  const text = new Element('span'); text.className = 'btn-text'; convert.appendChild(text);
  const loading = new Element('span'); loading.className = 'btn-loading'; loading.appendChild(new Element('span')); convert.appendChild(loading);
  const calls = [];
  const document = {
    createElement: tag => new Element(tag),
    querySelector: selector => elements[selector] || null,
    querySelectorAll: selector => {
      if (selector.startsWith('#field-list ')) return elements['#field-list'].querySelectorAll(selector.slice(12));
      if (selector === 'input, textarea, button') return Object.values(elements);
      return [];
    },
  };
  const context = vm.createContext({
    document, localStorage: { getItem: () => null, setItem() {} }, setTimeout() {}, clearTimeout() {}, console,
    window: { __TAURI__: { core: { invoke: async (name, args) => { calls.push({ name, args }); return [{ path: 'root', count: 1, is_matrix: false }]; } }, dialog: {} }, matchMedia: () => ({}), addEventListener() {} },
  });
  vm.runInContext(source, context);
  context.elements = elements;
  vm.runInContext(`inputPathEl=elements['#input-path']; outputPathEl=elements['#output-path']; btnConvert=elements['#btn-convert']; statusMsgEl=elements['#status-msg']; btnReveal=elements['#btn-reveal']; curlInputEl=elements['#curl-input']; jsonInputEl=elements['#json-input'];`, context);
  return { context, elements, calls, run: code => vm.runInContext(code, context) };
}

test('array keys remain literal text and checkbox values, selection survives language rerender', () => {
  const { context, elements, run } = setup();
  const key = 'root/summary <draft> "quoted" & 中文';
  context.fields = [{ path: key, count: 2, is_matrix: false }, { path: 'root/a~1b', count: 1, is_matrix: true }];
  run('showFieldSelector(fields)');
  const first = elements['#field-list'].children[0];
  assert.equal(first.children[0].value, key);
  assert.equal(first.children[2].textContent, key);
  assert.equal(first.children[2].children.length, 0);
  first.children[0].checked = false;
  run('currentLang="en"; showFieldSelector(fields)');
  assert.equal(elements['#field-list'].children[0].children[0].checked, false);
  assert.deepEqual(Array.from(run('getSelectedPaths()')), ['root/a~1b']);
});

test('paste and repeated parse preserve long numeric tokens without browser number conversion', async () => {
  const { elements, calls, run } = setup();
  const original = ' \n[{"id":123456789012345678901234567890,"nested":{"id":9007199254740993}}]\n ';
  elements['#json-input'].value = original;
  await run('activeTab="json"; parseJsonInput()');
  await run('parseJsonInput()');
  assert.equal(elements['#json-input'].value, original);
  assert.equal(calls.length, 2);
  assert.equal(calls[0].args.content, original.trim());
  assert.equal(calls[1].args.content, original.trim());
  run('invalidateParsedInput()');
  assert.equal(run('currentJsonContent'), '');
  assert.equal(run('rawFieldsOptions.length'), 0);
  assert.equal(elements['#btn-convert'].disabled, true);
});

test('pending work prevents repeated conversion and switching; cancellation restores usable UI', () => {
  const { elements, run } = setup();
  run('setLoading(true, "pending"); setLoading(true);');
  assert.equal(elements['#btn-convert'].disabled, true);
  assert.equal(elements['#json-input'].disabled, true);
  run('switchTab("json")');
  assert.equal(run('activeTab'), 'file');
  run('setLoading(false)');
  assert.equal(elements['#btn-convert'].disabled, true);
  run('setLoading(false)');
  assert.equal(elements['#json-input'].disabled, false);
  run('switchTab("json")');
  assert.equal(run('activeTab'), 'json');
});

test('strict CSP permits IPC and packaged assets without inline/eval sinks', () => {
  const config = JSON.parse(fs.readFileSync(path.join(root, 'src-tauri/tauri.conf.json'), 'utf8'));
  const csp = config.app.security.csp;
  assert.match(csp, /script-src 'self'/);
  assert.match(csp, /style-src 'self'/);
  assert.match(csp, /connect-src ipc: http:\/\/ipc.localhost/);
  assert.doesNotMatch(csp, /unsafe-inline|unsafe-eval/);
  assert.doesNotMatch(source, /innerHTML|outerHTML|insertAdjacentHTML|JSON\.parse|JSON\.stringify|\beval\s*\(/);
  assert.doesNotMatch(fs.readFileSync(path.join(root, 'src/index.html'), 'utf8'), /\sstyle=|\son\w+=/);
});
