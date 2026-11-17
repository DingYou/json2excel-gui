const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'src/main.js'), 'utf8');
const html = fs.readFileSync(path.join(root, 'src/index.html'), 'utf8');

class Element {
  constructor(tag = 'div', attrs = {}) {
    this.tagName = tag; this.attrs = { ...attrs }; this.children = []; this.style = {};
    this.value = ''; this.listeners = {}; this.disabled = false; this.textContent = '';
    this.classList = {
      add: name => { this.className = [...new Set([...(this.className || '').split(' ').filter(Boolean), name])].join(' '); },
      remove: name => { this.className = (this.className || '').split(' ').filter(value => value !== name).join(' '); },
      toggle: (name, enabled) => { if (enabled) this.classList.add(name); else this.classList.remove(name); },
    };
  }
  appendChild(child) { this.children.push(child); return child; }
  descendants() { return this.children.flatMap(child => [child, ...child.descendants()]); }
  replaceChildren(...children) {
    if (this.descendants().includes(this.ownerDocument?.activeElement)) this.ownerDocument.activeElement = null;
    this.children = children;
  }
  addEventListener(event, handler) { this.listeners[event] = handler; }
  setAttribute(key, value) { this.attrs[key] = value; }
  getAttribute(key) { return this.attrs[key] ?? null; }
  removeAttribute(key) { delete this.attrs[key]; }
  focus() { this.ownerDocument.activeElement = this; }
  querySelectorAll(selector) {
    const all = this.descendants();
    if (selector.includes("input[type='checkbox']")) return all.filter(el => el.type === 'checkbox' && (!selector.endsWith(':checked') || el.checked));
    if (selector.startsWith('.')) return all.filter(el => el.className?.split(' ').includes(selector.slice(1)));
    return all.filter(el => el.tagName === selector);
  }
  querySelector(selector) { return this.querySelectorAll(selector)[0] ?? null; }
  set innerHTML(_) { throw new Error('HTML injection sinks are not permitted'); }
}

function setup({ savedLang = null } = {}) {
  const elements = Object.fromEntries(Array.from(html.matchAll(/\bid="([^"]+)"/g), match => [`#${match[1]}`, new Element()]));
  const convert = elements['#btn-convert'];
  const text = new Element('span'); text.className = 'btn-text'; convert.appendChild(text);
  const loading = new Element('span'); loading.className = 'btn-loading'; loading.appendChild(new Element('span')); convert.appendChild(loading);
  const langText = new Element('span'); langText.className = 'lang-text'; langText.textContent = '中文'; elements['#btn-lang'].appendChild(langText);
  const translated = Array.from(html.matchAll(/\bdata-i18n="([^"]+)"/g), match => new Element('span', { 'data-i18n': match[1] }));
  const placeholders = Array.from(html.matchAll(/\bdata-i18n-placeholder="([^"]+)"/g), match => new Element('input', { 'data-i18n-placeholder': match[1] }));
  const calls = [];
  const storage = new Map(savedLang === null ? [] : [['json2excel-lang', savedLang]]);
  const callbacks = {};
  const document = {
    activeElement: null,
    documentElement: new Element('html', { lang: 'zh-CN' }),
    createElement: tag => { const element = new Element(tag); element.ownerDocument = document; return element; },
    querySelector: selector => elements[selector] || null,
    querySelectorAll: selector => {
      if (selector.startsWith('#field-list ')) return elements['#field-list'].querySelectorAll(selector.slice(12));
      if (selector === 'input, textarea, button') return Object.values(elements);
      if (selector === '[data-i18n]') return translated;
      if (selector === '[data-i18n-placeholder]') return placeholders;
      return [];
    },
  };
  for (const el of [...Object.values(elements), ...Object.values(elements).flatMap(el => el.descendants()), ...translated, ...placeholders, document.documentElement]) el.ownerDocument = document;
  const context = vm.createContext({
    document, localStorage: { getItem: key => storage.get(key) ?? null, setItem: (key, value) => storage.set(key, value) }, setTimeout() {}, clearTimeout() {}, console,
    window: {
      __TAURI__: { core: { invoke: async (name, args) => { calls.push({ name, args }); return [{ path: 'root', count: 1, is_matrix: false }]; } }, dialog: {} },
      matchMedia: () => ({ matches: false, addEventListener() {} }), addEventListener: (event, callback) => { callbacks[event] = callback; },
    },
  });
  vm.runInContext(source, context);
  callbacks.DOMContentLoaded();
  return { context, elements, calls, storage, document, translated, run: code => vm.runInContext(code, context) };
}

test('array keys remain literal text and checkbox values, selection survives rerender', () => {
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
  assert.doesNotMatch(html, /\sstyle=|\son\w+=/);
});

function assertLocale(app, lang) {
  const button = app.elements['#btn-lang'];
  assert.equal(app.run('currentLang'), lang);
  assert.equal(button.querySelector('.lang-text').textContent, lang === 'zh' ? '中文' : 'EN');
  assert.equal(button.getAttribute('data-tooltip'), lang === 'zh' ? '切换到英文' : 'Switch to Chinese');
  assert.equal(button.getAttribute('aria-label'), lang === 'zh' ? '当前语言：中文，切换到英文' : 'Current language: English. Switch to Chinese');
  assert.equal(app.document.documentElement.getAttribute('lang'), lang === 'zh' ? 'zh-CN' : 'en');
  assert.equal(app.translated.find(el => el.getAttribute('data-i18n') === 'tabFile').textContent, lang === 'zh' ? '文件' : 'File');
}

test('initial language markup and live boot show current language, target tooltip and accessible action', () => {
  assert.match(html, /id="btn-lang"[^>]*data-tooltip="切换到英文"[^>]*aria-label="当前语言：中文，切换到英文"/);
  assert.match(html, /class="lang-text">中文<\/span>/);
  assertLocale(setup(), 'zh');
});

test('actual click binding toggles in both directions and persists the selected current language', () => {
  const app = setup();
  const button = app.elements['#btn-lang'];
  button.focus(); button.listeners.click();
  assertLocale(app, 'en');
  assert.equal(app.storage.get('json2excel-lang'), 'en');
  assert.equal(app.document.activeElement, button);
  button.listeners.click();
  assertLocale(app, 'zh');
  assert.equal(app.storage.get('json2excel-lang'), 'zh');
  assert.equal(app.document.activeElement, button);
});

test('persisted boot restores matching labels/accessibility and invalid saved values fall back safely', () => {
  for (const lang of ['zh', 'en']) assertLocale(setup({ savedLang: lang }), lang);
  assertLocale(setup({ savedLang: 'invalid' }), 'zh');
});

test('locale change preserves checkbox identity/state/focus and pasted-text selection without rebuilding fields', () => {
  const app = setup();
  app.context.fields = [{ path: 'root/one', count: 2, is_matrix: false }, { path: 'root/two', count: 1, is_matrix: true }];
  app.run('rawFieldsOptions=fields; showFieldSelector(fields)');
  const checkboxes = app.elements['#field-list'].querySelectorAll("input[type='checkbox']");
  checkboxes[0].checked = false;
  checkboxes[1].focus();
  const input = app.elements['#json-input'];
  input.value = '[{"id":9007199254740993}]'; input.selectionStart = 7; input.selectionEnd = 23;
  app.run('toggleLang()');
  assertLocale(app, 'en');
  assert.equal(app.document.activeElement, checkboxes[1]);
  assert.equal(app.elements['#field-list'].querySelectorAll("input[type='checkbox']")[1], checkboxes[1]);
  assert.equal(checkboxes[0].checked, false);
  assert.deepEqual(Array.from(app.run('getSelectedPaths()')), ['root/two']);
  assert.match(app.elements['#field-list'].querySelectorAll('.field-meta')[1].textContent, /items, matrix/);
  const button = app.elements['#btn-lang']; button.focus(); button.listeners.click();
  assert.equal(app.document.activeElement, button);
  assert.equal(input.selectionStart, 7); assert.equal(input.selectionEnd, 23);
  assert.equal(input.value, '[{"id":9007199254740993}]');
  assert.equal(checkboxes[0].checked, false);
  assert.match(app.elements['#field-list'].querySelectorAll('.field-meta')[1].textContent, /项, 矩阵/);
});
