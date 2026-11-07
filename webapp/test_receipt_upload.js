'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ReceiptUpload = require('./static/receipt-upload.js');

const ROOT = __dirname;
const HTML = fs.readFileSync(path.join(ROOT, 'static', 'index.html'), 'utf8');
const APP = fs.readFileSync(path.join(ROOT, 'app.py'), 'utf8');
const JOB = '0123456789abcdef';

function extensionsIn(source, pattern) {
  const block = source.match(pattern);
  assert.ok(block, pattern);
  return [...block[1].matchAll(/['"](\.[a-z0-9]+)['"]/g)].map(match => match[1]).sort();
}

function inlineScript(html) {
  const parts = [];
  const re = /<script(\s[^>]*)?>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = re.exec(html))) {
    if (/\bsrc\s*=/.test(match[1] || '')) continue;
    parts.push(match[2]);
  }
  return parts.join('\n');
}

const PAGE = inlineScript(HTML);

function fileNamed(name, type, contents) {
  return new File([contents == null ? 'x' : contents], name, { type: type || '' });
}

function fileEntry(name, type) {
  const file = fileNamed(name, type || (String(name).toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'image/jpeg'));
  return {
    name,
    isFile: true,
    isDirectory: false,
    file(ok) { ok(file); }
  };
}

function brokenEntry(name) {
  return {
    name,
    isFile: true,
    isDirectory: false,
    file(ok, err) { err(new Error('denied')); }
  };
}

function directory(name, batches, calls) {
  return {
    name,
    isDirectory: true,
    isFile: false,
    createReader() {
      let index = 0;
      return {
        readEntries(success) {
          const batch = index < batches.length ? batches[index] : [];
          index += 1;
          if (calls) calls.push(batch.length);
          success(batch);
        }
      };
    }
  };
}

function files(count, prefix) {
  return Array.from({ length: count }, (_, index) => fileEntry(prefix + index + '.jpg'));
}

function transfer(items, loose) {
  return { types: ['Files'], items: items || [], files: loose || [] };
}

function jsonResult(status, body) {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body
  };
}

function sampleRow() {
  return {
    key: 'market#0',
    file: 'market.jpg',
    image: 'market.jpg',
    vendor: '한밭마트',
    date: '2026-08-05',
    total: 12500,
    items: 2,
    needs_check: true,
    confirmed: false,
    edited: false,
    excluded: false,
    note: '합계가 품목 합과 다릅니다',
    hint: ''
  };
}

async function flush() {
  await new Promise(resolve => setImmediate(resolve));
}

function createDom() {
  const elements = {};
  function makeEl(id) {
    const el = {
      id,
      style: {},
      disabled: false,
      value: '',
      checked: false,
      textContent: '',
      files: null,
      listeners: {},
      onclick: null,
      contains(node) { return node === el; },
      addEventListener(type, fn) {
        (el.listeners[type] ||= []).push(fn);
      },
      dispatchEvent(event) {
        const e = event || {};
        e.type = e.type || 'click';
        e.target = e.target || el;
        e.preventDefault = e.preventDefault || (() => { e.defaultPrevented = true; });
        e.stopPropagation = e.stopPropagation || (() => { e.stopped = true; });
        if (e.type === 'click' && el.onclick) el.onclick(e);
        for (const fn of el.listeners[e.type] || []) fn(e);
        return e;
      }
    };
    el.classList = {
      add(name) { el.classList._set.add(name); },
      remove(name) { el.classList._set.delete(name); },
      contains(name) { return el.classList._set.has(name); },
      _set: new Set()
    };
    Object.defineProperty(el, 'innerHTML', {
      get() { return el._html || ''; },
      set(html) { setRows(el, String(html)); }
    });
    elements[id] = el;
    return el;
  }
  function setRows(el, html) {
    el._html = html;
    if (el.id !== 'rows') return;
    for (const id of ['e_v', 'e_d', 'e_t', 'e_x']) delete elements[id];
    const re = /<input\b([^>]*)>/g;
    let match;
    while ((match = re.exec(html))) {
      const id = /id="([^"]+)"/.exec(match[1]);
      if (!id) continue;
      const input = makeEl(id[1]);
      const value = /value="([^"]*)"/.exec(match[1]);
      if (value) {
        input.value = value[1].replaceAll('&quot;', '"').replaceAll('&amp;', '&').replaceAll('&lt;', '<').replaceAll('&gt;', '>');
      }
      input.checked = /(?:^|\s)checked(?:\s|=|>|$)/.test(match[1]) || /\bchecked\b/.test(match[1]);
    }
  }
  ['drop', 'pick', 'files', 'go', 'err', 'step1', 'step2', 'step3', 'phase', 'pbar', 'pct', 'log', 'summary', 'toast', 'rows', 'dlLedger', 'dlImages', 'dlCsv', 'del', 'recover', 'recoverMsg', 'retryPoll', 'newUpload', 'relogin']
    .forEach(makeEl);
  elements.step2.classList.add('hide');
  elements.step3.classList.add('hide');
  elements.recover.classList.add('hide');
  elements.relogin.classList.add('hide');
  elements.go.disabled = true;
  elements.go.textContent = '정리 시작';
  elements.dlLedger.textContent = '지출현황 엑셀 받기';
  const buckets = { capture: {}, bubble: {} };
  const document = {
    addEventListener(type, fn, capture) {
      const bag = capture ? buckets.capture : buckets.bubble;
      (bag[type] ||= []).push(fn);
    },
    querySelector(selector) {
      return elements[String(selector || '').replace(/^#/, '')] || null;
    },
    dispatch(target, event) {
      const e = event || {};
      e.target = target;
      e.preventDefault = e.preventDefault || (() => { e.defaultPrevented = true; });
      e.stopPropagation = e.stopPropagation || (() => { e.stopped = true; });
      for (const fn of buckets.capture[e.type] || []) fn(e);
      if (e.stopped) return e;
      if (target && target !== document) {
        for (const fn of (target.listeners && target.listeners[e.type]) || []) fn(e);
      }
      if (e.stopped) return e;
      for (const fn of buckets.bubble[e.type] || []) fn(e);
      return e;
    }
  };
  return { document, elements };
}

function boot(options = {}) {
  const calls = [];
  const alerts = [];
  const timeouts = [];
  const intervals = [];
  const cleared = [];
  const store = new Map();
  if (options.lastJob) store.set('lastJob', options.lastJob);
  const routes = {
    handle: options.fetch || (async () => jsonResult(404, { detail: '없는 작업입니다.' }))
  };
  const { document, elements } = createDom();
  const location = {
    href: 'http://local/app',
    reloaded: false,
    reload() { this.reloaded = true; }
  };
  const sandbox = {
    console,
    document,
    location,
    localStorage: {
      getItem(key) { return store.has(key) ? store.get(key) : null; },
      setItem(key, value) { store.set(key, String(value)); },
      removeItem(key) { store.delete(key); }
    },
    alert(message) { alerts.push(String(message)); },
    confirm() { return options.confirmResult !== false; },
    fetch: async (url, opts) => {
      calls.push({
        url,
        method: (opts && opts.method) || 'GET',
        body: opts && opts.body,
        credentials: opts && opts.credentials
      });
      return routes.handle(url, opts || {});
    },
    FormData,
    File,
    setTimeout(fn, ms) {
      const item = { fn, ms: ms || 0 };
      timeouts.push(item);
      if (!ms) Promise.resolve().then(() => fn());
      return timeouts.length;
    },
    clearTimeout() {},
    setInterval(fn, ms) {
      const id = intervals.length + 1;
      intervals.push({ id, fn, ms });
      return id;
    },
    clearInterval(id) { cleared.push(id); }
  };
  if (options.withLibrary !== false) sandbox.ReceiptUpload = ReceiptUpload;
  vm.createContext(sandbox);
  vm.runInContext(PAGE, sandbox, { filename: 'index.html' });
  return { sandbox, document, elements, calls, alerts, timeouts, intervals, cleared, location, store, routes };
}

function postedNames(calls) {
  const post = calls.find(call => call.method === 'POST' && call.url === '/api/jobs');
  assert.ok(post, 'job upload was not sent');
  return post.body.getAll('files').map(file => file.name);
}

test('client extensions match the backend allow-list', () => {
  const backend = extensionsIn(APP, /ALLOWED_EXT = \{([^}]+)\}/);
  const browser = extensionsIn(fs.readFileSync(path.join(ROOT, 'static', 'receipt-upload.js'), 'utf8'), /ALLOWED_EXT = new Set\(\[([^\]]+)\]\)/);
  assert.deepEqual([...ReceiptUpload.ALLOWED_EXT].sort(), backend);
  assert.deepEqual(browser, backend);
  assert.equal(backend.includes('.gif') || backend.includes('.heic'), false);
});

test('page keeps job, review, export, and restore hooks', () => {
  assert.match(HTML, /src="\/static\/receipt-upload\.js"/);
  assert.equal(HTML.includes('dataTransfer.files'), false);
  assert.equal(HTML.includes('image/*'), false);
  const pickCss = HTML.match(/#pick\{([^}]+)\}/);
  assert.ok(pickCss);
  assert.equal(pickCss[1].replace(/\s/g, '').includes('display:none'), false);
  assert.match(pickCss[1], /position:\s*absolute/);
  for (const marker of [
    "localStorage.getItem('lastJob')",
    "localStorage.setItem('lastJob'",
    'setInterval(poll, 2000)',
    'download/ledger',
    'download/images',
    'download/json',
    'async function okRow',
    'async function saveRow',
    'function startEdit',
    'snapshotDataTransfer',
    'collectSnapshot',
    'submitJob',
    '한 번에 10개까지',
    '10MB',
    'id="retryPoll"',
    'id="newUpload"',
    'id="relogin"'
  ]) {
    assert.equal(HTML.includes(marker), true, marker);
  }
  assert.equal(HTML.includes('보관 기간이 지났을 수 있습니다'), false);
});

test('ordinary names follow the backend extensions', () => {
  const samples = [
    ['a.jpg', true], ['a.JPEG', true], ['a.jpeg', true], ['a.PNG', true], ['a.webp', true],
    ['a.bmp', true], ['a.tif', true], ['a.TIFF', true], ['a.PDF', true], ['폴더/하위/영수증.tif', true],
    ['a.gif', false], ['a.heic', false], ['a.heif', false], ['a.docx', false], ['a.txt', false],
    ['README', false], ['.DS_Store', false]
  ];
  for (const [name, ok] of samples) {
    assert.equal(ReceiptUpload.isSupportedName(name), ok, name);
  }
  const mixed = ReceiptUpload.classifyFileList([
    fileNamed('ok.jpg', 'image/jpeg'),
    fileNamed('no.gif', 'image/gif'),
    fileNamed('sheet.xlsx', 'application/vnd.ms-excel')
  ]);
  assert.deepEqual(mixed.files.map(file => file.name), ['ok.jpg']);
  assert.equal(mixed.unsupported.length, 2);
});

test('empty, unsupported, and folder placeholders get Korean feedback', () => {
  const empty = ReceiptUpload.describeSelection(ReceiptUpload.classifyFileList([]));
  assert.equal(empty.errorText, ReceiptUpload.MESSAGES.empty);
  assert.equal(empty.canStart, false);

  const bad = [1, 2, 3, 4, 5, 6].map(n => ({ name: '메모' + n + '.txt', type: 'text/plain', size: 3 }));
  bad.push({ name: '나쁜\n이름.doc', type: 'text/plain', size: 3 });
  const unsupported = ReceiptUpload.describeSelection(ReceiptUpload.classifyFileList(bad));
  assert.match(unsupported.errorText, /^지원하지 않는 형식입니다:/);
  assert.match(unsupported.errorText, /외 3개/);
  assert.equal(unsupported.errorText.includes('\n'), false);
  assert.equal(unsupported.canStart, false);

  const folder = ReceiptUpload.describeSelection(ReceiptUpload.classifyFileList([
    { name: '사진폴더', type: '', size: 4096 }
  ]));
  assert.equal(folder.errorText, ReceiptUpload.MESSAGES.folderUnreadable);

  const kept = ReceiptUpload.describeSelection(ReceiptUpload.classifyFileList([
    fileNamed('a.jpg', 'image/jpeg'), fileNamed('b.png', 'image/png'), { name: 'note.txt', type: 'text/plain', size: 1 }
  ]));
  assert.equal(kept.fileText, '선택한 파일 2개 · 지원하지 않는 파일 1개는 제외했습니다');
  assert.equal(kept.errorText, '');
  assert.equal(kept.canStart, true);
});

test('nested folders keep every readEntries batch, including short batches after 100', async () => {
  const rootCalls = [];
  const innerCalls = [];
  const inner = directory('inner', [files(100, 'i'), files(5, 'j')], innerCalls);
  const root = directory('root', [files(100, 'a'), [inner], files(50, 'b')], rootCalls);
  const flatCalls = [];
  const flat = directory('flat', [files(100, 'f'), files(100, 'g'), files(40, 'h')], flatCalls);
  const deep = directory('top', [[
    directory('mid', [[directory('leaf', [[fileEntry('deep.webp')]])]])
  ]]);

  const nested = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => root }
  ])));
  assert.deepEqual(rootCalls, [100, 1, 50, 0]);
  assert.deepEqual(innerCalls, [100, 5, 0]);
  assert.equal(nested.files.length, 255);
  assert.equal(nested.files[0].name, 'a0.jpg');
  assert.equal(nested.files.some(file => file.name === 'b49.jpg'), true);
  assert.equal(nested.files.some(file => file.name === 'j4.jpg'), true);
  assert.equal(nested.files.some(file => file.name === 'inner'), false);

  const many = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => flat }
  ])));
  assert.deepEqual(flatCalls, [100, 100, 40, 0]);
  assert.equal(many.files.length, 240);

  const three = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => deep }
  ])));
  assert.deepEqual(three.files.map(file => file.name), ['deep.webp']);
});

test('snapshot is taken before the live transfer is cleared, without duplicating files', async () => {
  let reads = 0;
  const entry = fileEntry('once.jpg');
  const loose = fileNamed('once.jpg', 'image/jpeg');
  const items = [{
    kind: 'file',
    webkitGetAsEntry() { reads += 1; return entry; },
    getAsFile() { return loose; }
  }];
  const files = [loose];
  const snap = ReceiptUpload.snapshotDataTransfer({ types: ['Files'], items, files });
  items.length = 0;
  files.length = 0;
  const result = await ReceiptUpload.collectSnapshot(snap);
  assert.equal(reads, 1);
  assert.deepEqual(result.files.map(file => file.name), ['once.jpg']);
});

test('fallback paths: missing entry API, null entry, thrown entry, and empty directory', async () => {
  const png = fileNamed('c.png', 'image/png');
  const fromFiles = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer({
    files: [png, fileNamed('skip.txt', 'text/plain')]
  }));
  assert.deepEqual(fromFiles.files.map(file => file.name), ['c.png']);

  const nullEntry = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer(
    [{ kind: 'file', webkitGetAsEntry: () => null, getAsFile: () => null }],
    [png]
  )));
  assert.deepEqual(nullEntry.files.map(file => file.name), ['c.png']);
  assert.equal(nullEntry.problems.length, 0);

  const thrown = fileNamed('rescued.jpg', 'image/jpeg');
  const rescued = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry() { throw new Error('consumed'); }, getAsFile: () => thrown }
  ])));
  assert.deepEqual(rescued.files.map(file => file.name), ['rescued.jpg']);
  assert.equal(rescued.problems.length, 0);

  const unprefixed = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', getAsEntry: () => fileEntry('plain.pdf') }
  ])));
  assert.deepEqual(unprefixed.files.map(file => file.name), ['plain.pdf']);

  let stringCalled = false;
  const ignored = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'string', webkitGetAsEntry() { stringCalled = true; return null; } },
    { kind: 'file', webkitGetAsEntry: () => fileEntry('keep.tif') }
  ])));
  assert.equal(stringCalled, false);
  assert.deepEqual(ignored.files.map(file => file.name), ['keep.tif']);

  const placeholder = { name: '영수증', type: '', size: 0 };
  const emptyDir = directory('영수증', []);
  const empty = ReceiptUpload.describeSelection(await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer(
    [{ kind: 'file', webkitGetAsEntry: () => emptyDir, getAsFile: () => placeholder }],
    [placeholder]
  ))));
  assert.equal(empty.errorText, ReceiptUpload.MESSAGES.empty);
  assert.equal(empty.canStart, false);

  const itemsThrow = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer({
    get items() { throw new Error('unsupported'); },
    files: [png]
  }));
  assert.deepEqual(itemsThrow.files.map(file => file.name), ['c.png']);

  const none = ReceiptUpload.describeSelection(await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(null)));
  assert.equal(none.errorText, ReceiptUpload.MESSAGES.empty);
});

test('a directory reader error keeps later files', async () => {
  const brokenDir = {
    name: 'broken',
    isDirectory: true,
    isFile: false,
    createReader() {
      return { readEntries(ok, err) { err(new Error('disk')); } };
    }
  };
  const result = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => brokenDir },
    { kind: 'file', webkitGetAsEntry: () => brokenEntry('bad.jpg') },
    { kind: 'file', webkitGetAsEntry: () => fileEntry('good.jpg') }
  ])));
  const view = ReceiptUpload.describeSelection(result);
  assert.deepEqual(view.files.map(file => file.name), ['good.jpg']);
  assert.equal(view.errorText, ReceiptUpload.partialReadText(2));
  assert.equal(view.canStart, true);

  const only = ReceiptUpload.describeSelection(await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => brokenDir }
  ]))));
  assert.equal(only.errorText, ReceiptUpload.MESSAGES.readError);
  assert.equal(only.canStart, false);
});

test('the batch cap returns files already read and reports the folder', async () => {
  let n = 0;
  const endless = {
    name: 'endless',
    isDirectory: true,
    isFile: false,
    createReader() {
      return {
        readEntries(success) {
          success([fileEntry('n' + (n++) + '.jpg')]);
        }
      };
    }
  };
  const result = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => endless }
  ])), { maxBatches: 3 });
  const view = ReceiptUpload.describeSelection(result);
  assert.deepEqual(view.files.map(file => file.name), ['n0.jpg', 'n1.jpg', 'n2.jpg']);
  assert.equal(view.canStart, true);
  assert.match(view.errorText, /일부 파일 1개/);
});

test('a reader that never calls back fails instead of leaving the UI waiting', { timeout: 2000 }, async () => {
  const hung = {
    name: 'hung',
    isDirectory: true,
    isFile: false,
    createReader() { return { readEntries() {} }; }
  };
  const hungFile = {
    name: 'wait.jpg',
    isFile: true,
    isDirectory: false,
    file() {}
  };
  const dirOnly = ReceiptUpload.describeSelection(await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => hung }
  ])), { timeoutMs: 30 }));
  assert.equal(dirOnly.errorText, ReceiptUpload.MESSAGES.readError);

  const mixed = ReceiptUpload.describeSelection(await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => hungFile },
    { kind: 'file', webkitGetAsEntry: () => fileEntry('ready.png') }
  ])), { timeoutMs: 30 }));
  assert.deepEqual(mixed.files.map(file => file.name), ['ready.png']);
  assert.match(mixed.errorText, /일부 파일 1개/);
});

test('callback data wins when a reader also returns a promise', async () => {
  let n = 0;
  const entry = {
    name: 'mixed',
    isDirectory: true,
    isFile: false,
    createReader() {
      return {
        readEntries(success) {
          n += 1;
          if (n > 1) return Promise.resolve([]);
          success([fileEntry('from-callback.jpg')]);
          return Promise.resolve([fileEntry('from-promise.jpg')]);
        }
      };
    }
  };
  const result = await ReceiptUpload.collectSnapshot(ReceiptUpload.snapshotDataTransfer(transfer([
    { kind: 'file', webkitGetAsEntry: () => entry }
  ])));
  assert.deepEqual(result.files.map(file => file.name), ['from-callback.jpg']);
});

test('selection gate keeps only the latest token', () => {
  const gate = ReceiptUpload.createSelectionGate();
  const first = gate.start();
  const second = gate.start();
  assert.equal(gate.isCurrent(first), false);
  assert.equal(gate.isCurrent(second), true);
});

test('upload success, server errors, network errors, and empty input', async () => {
  let calls = 0;
  await assert.rejects(
    () => ReceiptUpload.submitJob([], () => { calls += 1; }),
    error => error.code === 'empty' && error.message === ReceiptUpload.MESSAGES.uploadEmpty
  );
  assert.equal(calls, 0);

  const files = [fileNamed('영수증.jpg', 'image/jpeg'), fileNamed('b.pdf', 'application/pdf')];
  const ok = await ReceiptUpload.submitJob(files, async (url, opts) => {
    assert.equal(url, '/api/jobs');
    assert.equal(opts.method, 'POST');
    assert.equal(opts.credentials, 'same-origin');
    assert.deepEqual(opts.body.getAll('files').map(file => file.name), ['영수증.jpg', 'b.pdf']);
    return jsonResult(200, { job_id: JOB, file_count: 2, skipped: [] });
  });
  assert.equal(ok.job_id, JOB);

  await assert.rejects(
    () => ReceiptUpload.submitJob(files, async () => jsonResult(400, { detail: '한 번에 10개까지 올릴 수 있습니다.' })),
    error => error.message === '한 번에 10개까지 올릴 수 있습니다.' && error.status === 400
  );
  await assert.rejects(
    () => ReceiptUpload.submitJob(files, async () => jsonResult(422, { detail: [{ msg: 'field required' }, { msg: '다시 선택' }] })),
    /field required 다시 선택/
  );
  await assert.rejects(
    () => ReceiptUpload.submitJob(files, async () => ({ ok: false, status: 500, json: async () => { throw new Error('<html>'); } })),
    error => error.message === ReceiptUpload.MESSAGES.uploadFailed && error.code === 'http'
  );
  await assert.rejects(
    () => ReceiptUpload.submitJob(files, async () => { throw new Error('offline'); }),
    error => error.code === 'network' && error.message === ReceiptUpload.MESSAGES.uploadNetwork
  );
  await assert.rejects(
    () => ReceiptUpload.submitJob(files, async () => jsonResult(200, { file_count: 1 })),
    error => error.code === 'bad-response'
  );
  const long = '가'.repeat(500);
  await assert.rejects(
    () => ReceiptUpload.submitJob(files, async () => jsonResult(400, { detail: long })),
    error => error.message.length === 300
  );
});

test('picker selection uploads, then review, excel, edit, and delete still run', async () => {
  const patches = [];
  const page = boot({
    fetch: async (url, opts) => {
      if (url === '/api/jobs' && opts.method === 'POST') return jsonResult(200, { job_id: JOB, file_count: 2, skipped: ['old.heic'] });
      if (url === '/api/jobs/' + JOB && opts.method === 'DELETE') return jsonResult(200, { deleted: true });
      if (url === '/api/jobs/' + JOB) return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 2, skipped: ['old.heic'], log_tail: ['ok'] });
      if (url === '/api/jobs/' + JOB + '/receipts' && opts.method !== 'PATCH') {
        return jsonResult(200, { rows: [sampleRow()], ledger_dirty: false });
      }
      if (opts.method === 'PATCH') {
        patches.push({ url, body: JSON.parse(opts.body) });
        return jsonResult(200, { row: { ...sampleRow(), ...JSON.parse(opts.body), needs_check: true } });
      }
      return jsonResult(404, {});
    }
  });
  page.elements.pick.files = [fileNamed('a.jpg', 'image/jpeg'), fileNamed('b.pdf', 'application/pdf'), fileNamed('c.txt', 'text/plain')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  assert.equal(page.elements.files.textContent, '선택한 파일 2개 · 지원하지 않는 파일 1개는 제외했습니다');
  assert.equal(page.elements.go.disabled, false);
  assert.equal(page.elements.err.textContent, '');

  await page.elements.go.onclick();
  await flush();
  assert.deepEqual(postedNames(page.calls), ['a.jpg', 'b.pdf']);
  assert.equal(page.calls.find(call => call.method === 'POST').credentials, 'same-origin');
  assert.equal(page.store.get('lastJob'), JOB);
  assert.equal(page.elements.step1.classList.contains('hide'), true);
  assert.match(page.elements.rows.innerHTML, /한밭마트/);
  assert.match(page.elements.rows.innerHTML, /확인 필요/);
  assert.match(page.elements.rows.innerHTML, new RegExp('/api/jobs/' + JOB + '/image/market.jpg'));
  assert.match(page.elements.summary.innerHTML, /금액 확인이 필요합니다/);
  assert.match(page.elements.pct.textContent, /제외 1개/);

  page.elements.dlLedger.onclick();
  page.elements.dlImages.onclick();
  page.elements.dlCsv.onclick();
  assert.equal(page.location.href, '/api/jobs/' + JOB + '/download/json');
  page.elements.dlLedger.onclick();
  assert.equal(page.location.href, '/api/jobs/' + JOB + '/download/ledger');

  page.sandbox.startEdit(0);
  assert.equal(page.elements.e_v.value, '한밭마트');
  const before = page.calls.length;
  await page.sandbox.saveRow(0);
  assert.equal(page.calls.length, before);
  page.sandbox.startEdit(0);
  page.elements.e_v.value = '다른상점';
  page.elements.e_d.value = '2026-08-06';
  page.elements.e_t.value = '13000';
  await page.sandbox.saveRow(0);
  assert.deepEqual(patches.at(-1).body, { vendor: '다른상점', date: '2026-08-06', total: 13000 });
  assert.equal(patches.at(-1).url, '/api/jobs/' + JOB + '/receipts/market%230');
  assert.equal(page.elements.toast.textContent, '저장했습니다. 엑셀은 받을 때 고친 값으로 만들어집니다.');

  await page.sandbox.okRow(0);
  assert.deepEqual(patches.at(-1).body, { confirm: true });
  assert.equal(page.elements.toast.textContent, '확인했습니다. 엑셀은 받을 때 반영됩니다.');

  await page.elements.del.onclick();
  assert.equal(page.calls.some(call => call.method === 'DELETE' && call.url === '/api/jobs/' + JOB), true);
  assert.equal(page.store.has('lastJob'), false);
  assert.equal(page.location.reloaded, true);
});

test('a failed upload does not start polling or erase a stored job', async () => {
  const page = boot();
  page.store.set('lastJob', 'previous-job');
  page.elements.pick.files = [fileNamed('a.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  page.routes.handle = async () => jsonResult(400, { detail: '한 번에 10개까지 올릴 수 있습니다.' });
  await page.elements.go.onclick();
  assert.equal(page.elements.err.textContent, '한 번에 10개까지 올릴 수 있습니다.');
  assert.equal(page.elements.go.disabled, false);
  assert.equal(page.elements.step1.classList.contains('hide'), false);
  assert.equal(page.store.get('lastJob'), 'previous-job');
  assert.equal(page.intervals.length, 0);

  page.routes.handle = async () => { throw new Error('offline'); };
  await page.elements.go.onclick();
  assert.equal(page.elements.err.textContent, ReceiptUpload.MESSAGES.uploadNetwork);
  assert.equal(page.elements.go.disabled, false);
  assert.equal(page.store.get('lastJob'), 'previous-job');
});

test('changing the selection during upload does not change that job', async () => {
  let release;
  let posted;
  const page = boot({
    fetch: async (url, opts) => {
      if (url === '/api/jobs' && opts.method === 'POST') {
        posted = opts;
        await new Promise(resolve => { release = resolve; });
        return jsonResult(200, { job_id: JOB, file_count: 1, skipped: [] });
      }
      if (url === '/api/jobs/' + JOB) return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] });
      if (url.endsWith('/receipts')) return jsonResult(200, { rows: [sampleRow()], ledger_dirty: false });
      return jsonResult(404, {});
    }
  });
  page.elements.pick.files = [fileNamed('first.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  const pending = page.elements.go.onclick();
  for (let i = 0; i < 10 && !release; i++) await flush();
  assert.ok(release);
  page.elements.pick.files = [fileNamed('second.png', 'image/png')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  assert.equal(page.elements.files.textContent, '선택한 파일 1개');
  release();
  await pending;
  await flush();
  assert.deepEqual(posted.body.getAll('files').map(file => file.name), ['first.jpg']);
  assert.equal(page.store.get('lastJob'), JOB);
  assert.match(page.elements.rows.innerHTML, /한밭마트/);
  assert.equal(page.elements.step3.classList.contains('hide'), false);
});

test('folder drop feeds the same upload, and a newer picker selection wins the race', async () => {
  const calls = [];
  const inner = directory('inner', [files(100, 'i'), files(1, 'j')], calls);
  const root = directory('root', [files(100, 'a'), [inner], files(20, 'b')]);
  let entryReads = 0;
  const item = {
    kind: 'file',
    webkitGetAsEntry() { entryReads += 1; return root; }
  };
  const page = boot({
    fetch: async (url, opts) => {
      if (opts.method === 'POST') return jsonResult(200, { job_id: JOB, file_count: 1, skipped: [] });
      if (url === '/api/jobs/' + JOB) return jsonResult(200, { state: 'queued', phase: '대기 중', progress: 0, file_count: 1, queue_ahead: 1, skipped: [], log_tail: [] });
      return jsonResult(404, {});
    }
  });
  const event = page.document.dispatch(page.elements.drop, {
    type: 'drop',
    dataTransfer: transfer([item], [{ name: 'root', type: '', size: 4096 }])
  });
  await flush();
  assert.equal(event.stopped, true);
  assert.equal(entryReads, 1);
  assert.equal(page.elements.files.textContent, '선택한 파일 221개');
  assert.equal(page.elements.go.disabled, false);
  const pending = page.elements.go.onclick();
  await pending;
  await flush();
  assert.equal(postedNames(page.calls).length, 221);
  assert.equal(page.elements.step2.classList.contains('hide'), false);
  assert.match(page.elements.phase.textContent, /대기 중/);
  assert.equal(page.cleared.includes(page.intervals[0].id), false);

  let releaseRead;
  const slow = {
    name: 'slow',
    isDirectory: true,
    isFile: false,
    createReader() {
      let n = 0;
      return {
        readEntries(ok) {
          n += 1;
          if (n === 1) {
            releaseRead = () => ok([fileEntry('late.jpg')]);
            return;
          }
          ok([]);
        }
      };
    }
  };
  const race = boot();
  race.document.dispatch(race.elements.drop, {
    type: 'drop',
    dataTransfer: transfer([{ kind: 'file', webkitGetAsEntry: () => slow }])
  });
  assert.equal(race.elements.files.textContent, ReceiptUpload.MESSAGES.reading);
  assert.equal(race.elements.go.disabled, true);
  await race.elements.go.onclick();
  assert.equal(race.calls.some(call => call.method === 'POST'), false);

  race.elements.pick.files = [fileNamed('now.png', 'image/png')];
  race.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  releaseRead();
  await flush();
  assert.equal(race.elements.files.textContent, '선택한 파일 1개');
  race.routes.handle = async (url, opts) => {
    if (opts.method === 'POST') return jsonResult(200, { job_id: JOB, file_count: 1, skipped: [] });
    return jsonResult(200, { state: 'error', phase: '오류', error: '처리에 실패했습니다.', progress: 0, file_count: 1, skipped: [], log_tail: [] });
  };
  await race.elements.go.onclick();
  await flush();
  assert.deepEqual(postedNames(race.calls), ['now.png']);
  assert.match(race.elements.phase.textContent, /오류: 처리에 실패했습니다/);
  assert.equal(race.elements.step1.classList.contains('hide'), true);
});

test('drops outside the box work, and drops do not reset a restored or finished job', async () => {
  const page = boot();
  const types = { 0: 'Files', length: 1 };
  page.document.dispatch(page.document, {
    type: 'drop',
    dataTransfer: { types, items: [], files: [fileNamed('outside.jpg', 'image/jpeg')] }
  });
  await flush();
  assert.equal(page.elements.files.textContent, '선택한 파일 1개');
  assert.equal(page.elements.go.disabled, false);

  page.elements.step1.classList.add('hide');
  page.document.dispatch(page.document, {
    type: 'drop',
    dataTransfer: transfer([], [fileNamed('ignored.png', 'image/png')])
  });
  await flush();
  assert.equal(page.elements.files.textContent, '선택한 파일 1개');

  const restored = boot({
    lastJob: JOB,
    fetch: async url => {
      if (url === '/api/jobs/' + JOB) return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] });
      if (url.endsWith('/receipts')) return jsonResult(200, { rows: [sampleRow()], ledger_dirty: true });
      return jsonResult(404, {});
    }
  });
  await flush();
  assert.match(restored.elements.rows.innerHTML, /한밭마트/);
  assert.equal(restored.elements.step1.classList.contains('hide'), true);
  restored.document.dispatch(restored.document, {
    type: 'drop',
    dataTransfer: transfer([], [fileNamed('late.jpg', 'image/jpeg')])
  });
  await flush();
  assert.match(restored.elements.rows.innerHTML, /한밭마트/);
  assert.equal(restored.store.get('lastJob'), JOB);
  restored.elements.dlLedger.onclick();
  assert.equal(restored.elements.dlLedger.textContent, '고친 값으로 엑셀 만드는 중…');
  assert.equal(restored.location.href, '/api/jobs/' + JOB + '/download/ledger');
  const pending = restored.timeouts.find(item => item.ms === 7000);
  pending.fn();
  assert.equal(restored.elements.dlLedger.disabled, false);
  assert.equal(restored.elements.dlLedger.textContent, '지출현황 엑셀 받기');

  const missing = boot({ lastJob: 'gone', fetch: async () => jsonResult(404, {}) });
  await flush();
  assert.equal(missing.store.has('lastJob'), false);
  assert.equal(missing.elements.step1.classList.contains('hide'), false);
  assert.equal(missing.elements.step2.classList.contains('hide'), true);
  assert.match(missing.elements.err.textContent, /이전에 보던 작업을 서버에서 찾지 못했습니다/);
  assert.equal(missing.intervals.length, 0);

  const queued = boot({
    lastJob: JOB,
    fetch: async () => jsonResult(200, { state: 'running', phase: '영수증 내용 읽는 중', progress: 40, file_count: 3, skipped: [], log_tail: ['[1/3]'] })
  });
  await flush();
  assert.equal(queued.elements.step2.classList.contains('hide'), false);
  assert.equal(queued.elements.phase.textContent, '영수증 내용 읽는 중');
  assert.equal(queued.elements.pbar.style.width, '40%');
  assert.equal(queued.intervals.length, 1);
});

test('Korean empty, unsupported, folder, and read messages show on the page', async () => {
  const page = boot();
  page.elements.pick.files = [];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  assert.equal(page.elements.err.textContent, ReceiptUpload.MESSAGES.empty);
  assert.equal(page.elements.go.disabled, true);

  page.elements.pick.files = [fileNamed('메모.txt', 'text/plain'), fileNamed('사진.heic', 'image/heic')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  assert.match(page.elements.err.textContent, /지원하지 않는 형식입니다: 메모.txt, 사진.heic/);
  assert.equal(page.elements.go.disabled, true);

  page.document.dispatch(page.document, {
    type: 'drop',
    dataTransfer: { types: ['Files'], items: [], files: [{ name: '사진폴더', type: '', size: 4096 }] }
  });
  await flush();
  assert.equal(page.elements.err.textContent, ReceiptUpload.MESSAGES.folderUnreadable);

  const broken = {
    name: 'broken',
    isDirectory: true,
    isFile: false,
    createReader() { return { readEntries(ok, err) { err(new Error('disk')); } }; }
  };
  page.document.dispatch(page.elements.drop, {
    type: 'drop',
    dataTransfer: transfer([{ kind: 'file', webkitGetAsEntry: () => broken }])
  });
  await flush();
  assert.equal(page.elements.err.textContent, ReceiptUpload.MESSAGES.readError);
  assert.equal(page.elements.go.disabled, true);
});

test('a missing upload script still restores the last job', async () => {
  const page = boot({
    withLibrary: false,
    lastJob: JOB,
    fetch: async url => {
      if (url === '/api/jobs/' + JOB) return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] });
      if (url.endsWith('/receipts')) return jsonResult(200, { rows: [sampleRow()], ledger_dirty: false });
      return jsonResult(404, {});
    }
  });
  await flush();
  assert.match(page.elements.err.textContent, /업로드 기능을 불러오지 못했습니다/);
  assert.match(page.elements.rows.innerHTML, /한밭마트/);
  assert.equal(page.store.get('lastJob'), JOB);
  page.elements.dlImages.onclick();
  assert.equal(page.location.href, '/api/jobs/' + JOB + '/download/images');

  const bare = boot({ withLibrary: false });
  assert.match(bare.elements.err.textContent, /업로드 기능을 불러오지 못했습니다/);
  assert.equal(bare.elements.go.disabled, true);
  assert.equal(bare.elements.step1.classList.contains('hide'), false);
});

test('delete asks first and a lost job poll explains that the job is gone', async () => {
  const declined = boot({ confirmResult: false });
  await declined.elements.del.onclick();
  assert.equal(declined.calls.length, 0);
  assert.equal(declined.location.reloaded, false);

  const page = boot({
    fetch: async (url, opts) => {
      if (opts.method === 'POST') return jsonResult(200, { job_id: JOB, file_count: 1, skipped: [] });
      return jsonResult(404, {});
    }
  });
  page.elements.pick.files = [fileNamed('a.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  await page.elements.go.onclick();
  await flush();
  assert.equal(page.elements.phase.textContent, '작업을 서버에서 찾지 못했습니다');
  assert.match(page.elements.recoverMsg.textContent, /복구할 수 없습니다/);
  assert.equal(page.elements.recover.classList.contains('hide'), false);
  assert.equal(page.store.has('lastJob'), false);
  assert.equal(page.cleared.length > 0, true);
  assert.equal(page.elements.phase.textContent.includes('보관 기간'), false);
});

function runningStatus() {
  return { state: 'running', phase: '사진 정리 중', progress: 1, file_count: 10, skipped: [], log_tail: [], queue_ahead: 0 };
}

test('a restored missing job stays on the upload form without a timer', async () => {
  const page = boot({ lastJob: JOB, fetch: async () => jsonResult(404, { detail: '없는 작업입니다.' }) });
  await flush();
  assert.equal(page.store.has('lastJob'), false);
  assert.equal(page.elements.step1.classList.contains('hide'), false);
  assert.equal(page.elements.step2.classList.contains('hide'), true);
  assert.match(page.elements.err.textContent, /이전에 보던 작업을 서버에서 찾지 못했습니다/);
  assert.equal(page.intervals.length, 0);
  assert.equal(page.elements.phase.textContent.includes('보관 기간'), false);
});

test('a job that disappears while polling can be replaced by a new upload', async () => {
  let mode = 'running';
  const nextJob = 'abcdef0123456789';
  const page = boot({
    lastJob: JOB,
    fetch: async (url, opts) => {
      if (opts.method === 'POST') return jsonResult(200, { job_id: nextJob, file_count: 1, skipped: [] });
      if (mode === '404') return jsonResult(404, { detail: '없는 작업입니다.' });
      if (String(url).endsWith('/receipts')) return jsonResult(200, { rows: [sampleRow()], ledger_dirty: false });
      if (mode === 'done') return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] });
      return jsonResult(200, runningStatus());
    }
  });
  await flush();
  assert.equal(page.elements.pbar.style.width, '1%');
  assert.match(page.elements.pct.textContent, /영수증 10개/);
  assert.equal(page.store.get('lastJob'), JOB);
  const intervalId = page.intervals.at(-1).id;
  mode = '404';
  await page.intervals.at(-1).fn();
  await flush();
  assert.equal(page.elements.phase.textContent, '작업을 서버에서 찾지 못했습니다');
  assert.match(page.elements.recoverMsg.textContent, /복구할 수 없습니다/);
  assert.equal(page.elements.pct.textContent, '서버에 남은 진행 정보가 없습니다.');
  assert.equal(page.elements.recover.classList.contains('hide'), false);
  assert.equal(page.store.has('lastJob'), false);
  assert.equal(page.cleared.includes(intervalId), true);
  assert.equal(page.elements.step2.classList.contains('hide'), false);
  page.elements.newUpload.onclick();
  assert.equal(page.elements.step1.classList.contains('hide'), false);
  assert.equal(page.elements.step2.classList.contains('hide'), true);
  mode = 'done';
  page.elements.pick.files = [fileNamed('again.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  await page.elements.go.onclick();
  await flush();
  assert.deepEqual(postedNames(page.calls), ['again.jpg']);
  assert.equal(page.store.get('lastJob'), nextJob);
  assert.match(page.elements.rows.innerHTML, /한밭마트/);
  assert.equal(page.elements.step3.classList.contains('hide'), false);
});

test('network and 500 failures keep the job, the timer, and the last progress', async () => {
  let mode = 'running';
  const page = boot({
    lastJob: JOB,
    fetch: async () => {
      if (mode === '500') return jsonResult(500, { detail: 'temporary' });
      if (mode === 'down') throw new Error('offline');
      return jsonResult(200, runningStatus());
    }
  });
  await flush();
  const intervalId = page.intervals.at(-1).id;
  mode = '500';
  await page.intervals.at(-1).fn();
  await flush();
  assert.equal(page.elements.phase.textContent, '잠시 후 다시 확인합니다');
  assert.match(page.elements.recoverMsg.textContent, /지우지 않았고/);
  assert.equal(page.elements.phase.textContent.includes('찾지 못했습니다'), false);
  assert.equal(page.store.get('lastJob'), JOB);
  assert.equal(page.cleared.includes(intervalId), false);
  assert.match(page.elements.pct.textContent, /영수증 10개/);
  mode = 'down';
  await page.intervals.at(-1).fn();
  await flush();
  assert.equal(page.elements.phase.textContent, '연결을 확인하지 못했습니다');
  assert.match(page.elements.recoverMsg.textContent, /네트워크/);
  assert.equal(page.store.get('lastJob'), JOB);
  assert.equal(page.cleared.includes(intervalId), false);
  mode = 'running';
  page.elements.retryPoll.onclick();
  await flush();
  assert.equal(page.elements.phase.textContent, '사진 정리 중');
  assert.equal(page.elements.recover.classList.contains('hide'), true);
  assert.equal(page.elements.pbar.style.width, '1%');
  assert.equal(page.store.get('lastJob'), JOB);
});

test('an auth failure keeps the job and asks for login', async () => {
  let mode = 'running';
  const page = boot({
    lastJob: JOB,
    fetch: async () => mode === 'running'
      ? jsonResult(200, runningStatus())
      : jsonResult(401, { detail: '접속 코드가 필요합니다.' })
  });
  await flush();
  mode = '401';
  await page.intervals.at(-1).fn();
  await flush();
  assert.equal(page.elements.phase.textContent, '로그인이 필요합니다');
  assert.match(page.elements.recoverMsg.textContent, /다시 로그인/);
  assert.equal(page.elements.phase.textContent.includes('보관 기간'), false);
  assert.equal(page.store.get('lastJob'), JOB);
  assert.equal(page.elements.relogin.classList.contains('hide'), false);
  assert.equal(page.elements.newUpload.classList.contains('hide'), true);
  assert.equal(page.cleared.includes(page.intervals.at(-1).id), true);
  page.elements.relogin.onclick();
  assert.equal(page.location.href, '/app');
});

test('a terminal processing error keeps the saved job id', async () => {
  const page = boot({
    lastJob: JOB,
    fetch: async () => jsonResult(200, {
      state: 'error', phase: '오류', error: '처리에 실패했습니다.', progress: 1, file_count: 10, skipped: [], log_tail: []
    })
  });
  await flush();
  assert.match(page.elements.phase.textContent, /오류: 처리에 실패했습니다/);
  assert.match(page.elements.recoverMsg.textContent, /서버의 파일은 지우지 않았습니다/);
  assert.equal(page.store.get('lastJob'), JOB);
  assert.equal(page.cleared.length > 0, true);
  page.elements.newUpload.onclick();
  assert.equal(page.elements.step1.classList.contains('hide'), false);
  assert.equal(page.store.get('lastJob'), JOB);
});

test('restoring while signed out does not open the progress screen', async () => {
  const page = boot({
    lastJob: JOB,
    fetch: async () => jsonResult(401, { detail: '접속 코드가 필요합니다.' })
  });
  await flush();
  assert.match(page.elements.err.textContent, /다시 로그인한 뒤/);
  assert.equal(page.elements.step2.classList.contains('hide'), true);
  assert.equal(page.store.get('lastJob'), JOB);
  assert.equal(page.intervals.length, 0);
});

test('stale 404, 401, done, and result responses cannot replace the current job', async () => {
  const nextJob = 'abcdef0123456789';
  const waiting = [];
  let hold = false;
  const page = boot({
    lastJob: JOB,
    fetch(url, opts) {
      const method = (opts && opts.method) || 'GET';
      if (method === 'POST') return jsonResult(200, { job_id: nextJob, file_count: 1, skipped: [] });
      if (hold) return new Promise(resolve => waiting.push({ url: String(url), resolve }));
      return jsonResult(200, runningStatus());
    }
  });
  await flush();
  hold = true;
  const statusPoll = page.sandbox.poll();
  const resultRows = [{ ...sampleRow(), vendor: '옛마트' }];
  await flush();
  page.elements.newUpload.onclick();
  page.elements.pick.files = [fileNamed('again.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  const started = page.elements.go.onclick();
  await flush();
  assert.equal(page.store.get('lastJob'), nextJob);
  const newInterval = page.intervals.at(-1).id;
  const oldStatus = waiting.find(item => item.url.endsWith('/' + JOB));
  oldStatus.resolve(jsonResult(404, { detail: '없는 작업입니다.' }));
  await statusPoll;
  await flush();
  assert.equal(page.store.get('lastJob'), nextJob);
  assert.equal(page.cleared.includes(newInterval), false);
  assert.equal(page.elements.phase.textContent.includes('찾지 못했습니다'), false);
  assert.equal(page.elements.relogin.classList.contains('hide'), true);

  const oldAuth = page.sandbox.poll();
  await flush();
  const authCall = waiting.at(-1);
  page.elements.newUpload.onclick();
  page.elements.pick.files = [fileNamed('third.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  const restarted = page.elements.go.onclick();
  await flush();
  authCall.resolve(jsonResult(401, { detail: '접속 코드가 필요합니다.' }));
  await oldAuth;
  await flush();
  assert.equal(page.store.get('lastJob'), nextJob);
  assert.equal(page.elements.phase.textContent.includes('로그인'), false);

  const donePoll = page.sandbox.poll();
  await flush();
  const doneCall = waiting.filter(item => item.url.endsWith('/' + nextJob)).at(-1);
  doneCall.resolve(jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] }));
  await flush();
  const receiptCall = waiting.find(item => item.url.includes('/receipts'));
  page.elements.newUpload.onclick();
  page.store.set('lastJob', nextJob);
  receiptCall.resolve(jsonResult(200, { rows: resultRows, ledger_dirty: false }));
  await donePoll;
  await flush();
  assert.equal(page.elements.rows.innerHTML.includes('옛마트'), false);
  assert.equal(page.elements.step3.classList.contains('hide'), true);
  assert.equal(page.store.get('lastJob'), nextJob);
  assert.equal(page.elements.step1.classList.contains('hide'), false);
});

test('an older poll cannot clear or restore lastJob around a newer poll', async () => {
  const waiting = [];
  let hold = false;
  const page = boot({
    lastJob: JOB,
    fetch() {
      if (hold) return new Promise(resolve => waiting.push(resolve));
      return jsonResult(200, runningStatus());
    }
  });
  await flush();
  hold = true;
  const older = page.sandbox.poll();
  const newer = page.sandbox.poll();
  await flush();
  waiting[1](jsonResult(200, { ...runningStatus(), progress: 40 }));
  await newer;
  await flush();
  assert.equal(page.elements.pbar.style.width, '40%');
  waiting[0](jsonResult(404, { detail: '없는 작업입니다.' }));
  await older;
  await flush();
  assert.equal(page.store.get('lastJob'), JOB);
  assert.equal(page.elements.pbar.style.width, '40%');
  assert.equal(page.elements.phase.textContent.includes('찾지 못했습니다'), false);

  const lateProgress = page.sandbox.poll();
  const lateMissing = page.sandbox.poll();
  await flush();
  waiting.at(-1)(jsonResult(404, { detail: '없는 작업입니다.' }));
  await lateMissing;
  await flush();
  assert.equal(page.store.has('lastJob'), false);
  waiting.at(-2)(jsonResult(200, runningStatus()));
  await lateProgress;
  await flush();
  assert.equal(page.store.has('lastJob'), false);
  assert.equal(page.elements.phase.textContent, '작업을 서버에서 찾지 못했습니다');
});

test('a slow restore 404 does not clear the job started meanwhile', async () => {
  const nextJob = 'abcdef0123456789';
  let restore = null;
  let sawRestore = false;
  const page = boot({
    lastJob: JOB,
    fetch(url, opts) {
      const method = (opts && opts.method) || 'GET';
      if (method === 'POST') return jsonResult(200, { job_id: nextJob, file_count: 1, skipped: [] });
      if (!sawRestore) {
        sawRestore = true;
        return new Promise(resolve => { restore = resolve; });
      }
      return jsonResult(200, { ...runningStatus(), file_count: 2, progress: 3 });
    }
  });
  for (let i = 0; i < 5 && !restore; i++) await flush();
  assert.ok(restore);
  page.elements.pick.files = [fileNamed('again.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  const started = page.elements.go.onclick();
  await flush();
  assert.equal(page.store.get('lastJob'), nextJob);
  restore(jsonResult(404, { detail: '없는 작업입니다.' }));
  await started;
  await flush();
  assert.equal(page.store.get('lastJob'), nextJob);
  assert.equal(page.elements.step1.classList.contains('hide'), true);
  assert.match(page.elements.pct.textContent, /영수증 2개/);
});

test('result 401 keeps the job and asks for login, while 500 stays retryable', async () => {
  const auth = boot({
    lastJob: JOB,
    fetch(url) {
      if (String(url).endsWith('/receipts')) return jsonResult(401, { detail: '접속 코드가 필요합니다.' });
      return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] });
    }
  });
  await flush();
  assert.equal(auth.elements.phase.textContent, '로그인이 필요합니다');
  assert.match(auth.elements.recoverMsg.textContent, /다시 로그인/);
  assert.equal(auth.elements.relogin.classList.contains('hide'), false);
  assert.equal(auth.elements.newUpload.classList.contains('hide'), true);
  assert.equal(auth.store.get('lastJob'), JOB);
  assert.equal(auth.elements.step3.classList.contains('hide'), true);
  assert.equal(auth.elements.rows.innerHTML.includes('한밭마트'), false);

  let receipts = '500';
  const retry = boot({
    lastJob: JOB,
    fetch(url) {
      if (String(url).endsWith('/receipts')) {
        if (receipts === '500') return jsonResult(500, { detail: 'temporary' });
        return jsonResult(200, { rows: [sampleRow()], ledger_dirty: false });
      }
      return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] });
    }
  });
  await flush();
  assert.equal(retry.elements.phase.textContent, '잠시 후 다시 확인합니다');
  assert.match(retry.elements.recoverMsg.textContent, /지우지 않았습니다/);
  assert.equal(retry.elements.phase.textContent.includes('로그인'), false);
  assert.equal(retry.elements.relogin.classList.contains('hide'), true);
  assert.equal(retry.store.get('lastJob'), JOB);
  receipts = '200';
  retry.elements.retryPoll.onclick();
  await flush();
  assert.match(retry.elements.rows.innerHTML, /한밭마트/);
  assert.equal(retry.elements.step3.classList.contains('hide'), false);
  assert.equal(retry.store.get('lastJob'), JOB);
});

test('polls slower than the 2 second interval still render progress and done', async () => {
  const waiting = [];
  let hold = false;
  const page = boot({
    lastJob: JOB,
    fetch(url) {
      if (hold || String(url).endsWith('/receipts')) {
        return new Promise(resolve => waiting.push({ url: String(url), resolve }));
      }
      return jsonResult(200, runningStatus());
    }
  });
  await flush();
  hold = true;
  const first = page.intervals.at(-1).fn();
  await flush();
  const second = page.intervals.at(-1).fn();
  await flush();
  assert.equal(waiting.length, 2);
  waiting[0].resolve(jsonResult(200, { ...runningStatus(), progress: 35 }));
  await first;
  await flush();
  assert.equal(page.elements.pbar.style.width, '35%');
  assert.equal(page.store.get('lastJob'), JOB);
  waiting[1].resolve(jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 10, skipped: [], log_tail: [] }));
  await second;
  await flush();
  const receipt = waiting.find(item => item.url.includes('/receipts'));
  assert.ok(receipt);
  receipt.resolve(jsonResult(200, { rows: [sampleRow()], ledger_dirty: false }));
  await flush();
  assert.match(page.elements.rows.innerHTML, /한밭마트/);
  assert.equal(page.elements.step3.classList.contains('hide'), false);
  assert.equal(page.store.get('lastJob'), JOB);
});

test('a timed out poll keeps the job, ignores its late body, and a retry can finish', async () => {
  const waiting = [];
  let mode = 'fast';
  const nextJob = 'abcdef0123456789';
  const page = boot({
    lastJob: JOB,
    fetch(url, opts) {
      const method = (opts && opts.method) || 'GET';
      if (method === 'POST') return jsonResult(200, { job_id: nextJob, file_count: 1, skipped: [] });
      if (mode === 'hold') return new Promise(resolve => waiting.push({ url: String(url), resolve }));
      if (String(url).endsWith('/receipts')) return jsonResult(200, { rows: [sampleRow()], ledger_dirty: false });
      if (mode === 'done') return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] });
      return jsonResult(200, { ...runningStatus(), progress: 55 });
    }
  });
  await flush();
  mode = 'hold';
  const slow = page.sandbox.poll();
  await flush();
  const timeout = page.timeouts.filter(item => item.ms === 15000).at(-1);
  assert.ok(timeout);
  timeout.fn();
  assert.equal(page.elements.phase.textContent, '잠시 후 다시 확인합니다');
  assert.match(page.elements.recoverMsg.textContent, /응답이 늦습니다/);
  assert.equal(page.store.get('lastJob'), JOB);
  waiting[0].resolve(jsonResult(404, { detail: '없는 작업입니다.' }));
  await slow;
  await flush();
  assert.equal(page.store.get('lastJob'), JOB);
  assert.equal(page.elements.phase.textContent.includes('찾지 못했습니다'), false);
  mode = 'fast';
  page.elements.retryPoll.onclick();
  await flush();
  assert.equal(page.elements.pbar.style.width, '55%');
  assert.equal(page.store.get('lastJob'), JOB);

  mode = 'hold';
  const oldJobPoll = page.sandbox.poll();
  await flush();
  const oldCall = waiting.at(-1);
  mode = 'done';
  page.elements.newUpload.onclick();
  page.elements.pick.files = [fileNamed('again.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  const started = page.elements.go.onclick();
  await started;
  await flush();
  oldCall.resolve(jsonResult(404, { detail: '없는 작업입니다.' }));
  await oldJobPoll;
  await flush();
  assert.equal(page.store.get('lastJob'), nextJob);
  assert.match(page.elements.rows.innerHTML, /한밭마트/);
  assert.equal(page.elements.phase.textContent.includes('찾지 못했습니다'), false);
});

test('patch failure alerts and does not pretend the row was saved', async () => {
  const page = boot({
    fetch: async (url, opts) => {
      if (opts.method === 'POST') return jsonResult(200, { job_id: JOB, file_count: 1, skipped: [] });
      if (opts.method === 'PATCH') return jsonResult(400, { detail: '상호는 1~60자로 적어 주세요.' });
      if (url === '/api/jobs/' + JOB) return jsonResult(200, { state: 'done', phase: '완료', progress: 100, file_count: 1, skipped: [], log_tail: [] });
      if (url.endsWith('/receipts')) return jsonResult(200, { rows: [sampleRow()], ledger_dirty: false });
      return jsonResult(404, {});
    }
  });
  page.elements.pick.files = [fileNamed('a.jpg', 'image/jpeg')];
  page.elements.pick.dispatchEvent({ type: 'change' });
  await flush();
  await page.elements.go.onclick();
  await flush();
  await page.sandbox.okRow(0);
  assert.deepEqual(page.alerts, ['상호는 1~60자로 적어 주세요.']);
  assert.match(page.elements.rows.innerHTML, /확인 필요/);
});
