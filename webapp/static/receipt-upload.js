/* 영수증 파일 선택.
   폴더를 떨어뜨리면 dataTransfer.files 에는 폴더 껍데기만 있거나 목록이 비어 있다.
   안의 파일은 webkitGetAsEntry 로만 보이고, readEntries 는 한 번에 최대 100개라
   빈 배열이 나올 때까지 다시 읽어야 한다. items 는 drop 처리가 끝나기 전에 복사한다. */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.ReceiptUpload = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var ALLOWED_EXT = new Set(['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tif', '.tiff', '.pdf']);
  var DEFAULT_TIMEOUT_MS = 10000;
  var DEFAULT_MAX_BATCHES = 1000;

  var MESSAGES = {
    reading: '파일을 확인하는 중…',
    empty: '올릴 수 있는 파일이 없습니다. 사진(JPG, PNG, WEBP, BMP, TIF)이나 PDF를 선택해 주세요.',
    readError: '파일을 읽지 못했습니다. 다시 끌어다 놓거나 선택해 주세요.',
    folderUnreadable: '폴더 안의 파일을 찾지 못했습니다. 폴더를 점선 칸에 다시 끌어다 놓거나, 사진(JPG, PNG, WEBP, BMP, TIF)이나 PDF를 직접 선택해 주세요.',
    uploadNetwork: '업로드에 실패했습니다. 네트워크를 확인하고 다시 시도해 주세요.',
    uploadFailed: '업로드에 실패했습니다.',
    uploadEmpty: '올릴 파일을 선택해 주세요.'
  };

  function toArray(list) {
    if (!list) return [];
    if (Array.isArray(list)) return list.slice();
    if (typeof list.length === 'number') {
      var out = [];
      for (var i = 0; i < list.length; i++) out.push(list[i]);
      return out;
    }
    return [];
  }

  function baseName(name) {
    var parts = String(name || '').split(/[/\\]/);
    return parts[parts.length - 1] || '';
  }

  function extensionOf(name) {
    var base = baseName(name);
    var dot = base.lastIndexOf('.');
    if (dot <= 0) return '';
    return base.slice(dot).toLowerCase();
  }

  function isSupportedName(name) {
    return ALLOWED_EXT.has(extensionOf(name));
  }

  function previewNames(items) {
    var names = items.map(function (item) {
      var raw = item && item.name ? item.name : item;
      return String(raw || '(이름 없음)').replace(/[\r\n\t]/g, ' ').slice(0, 80);
    });
    var shown = names.slice(0, 4);
    var extra = names.length - shown.length;
    return shown.join(', ') + (extra > 0 ? ' 외 ' + extra + '개' : '');
  }

  function unsupportedMessage(items) {
    return '지원하지 않는 형식입니다: ' + previewNames(items) + '. JPG, PNG, WEBP, BMP, TIF, PDF만 올릴 수 있습니다.';
  }

  function selectedText(count, skipped) {
    var text = '선택한 파일 ' + count + '개';
    if (skipped) text += ' · 지원하지 않는 파일 ' + skipped + '개는 제외했습니다';
    return text;
  }

  function partialReadText(count) {
    return '일부 파일 ' + count + '개를 읽지 못했습니다. 읽은 파일만 올릴 수 있습니다.';
  }

  function isFolderLike(file) {
    if (!file) return false;
    return !extensionOf(file.name) && !String(file.type || '');
  }

  function classifyFileList(files, problems) {
    var accepted = [];
    var unsupported = [];
    toArray(files).forEach(function (file) {
      if (!file) return;
      var name = file.name || '';
      if (isSupportedName(name)) accepted.push(file);
      else unsupported.push({ name: name || '(이름 없음)', folderLike: isFolderLike(file) });
    });
    return {
      files: accepted,
      unsupported: unsupported,
      problems: Array.isArray(problems) ? problems.slice() : []
    };
  }

  function describeSelection(result) {
    var files = result && result.files ? result.files : [];
    var unsupported = result && result.unsupported ? result.unsupported : [];
    var problems = result && result.problems ? result.problems : [];
    var readFail = 0;
    for (var i = 0; i < problems.length; i++) {
      if (!problems[i] || problems[i].reason === 'read') readFail += 1;
    }
    var allFolderLike = unsupported.length > 0;
    for (var j = 0; j < unsupported.length; j++) {
      if (!unsupported[j] || !unsupported[j].folderLike) allFolderLike = false;
    }
    var fileText = '';
    var errorText = '';
    if (files.length) fileText = selectedText(files.length, unsupported.length);
    if (!files.length && !unsupported.length && !readFail) errorText = MESSAGES.empty;
    else if (!files.length && readFail) errorText = MESSAGES.readError;
    else if (!files.length && allFolderLike) errorText = MESSAGES.folderUnreadable;
    else if (!files.length && unsupported.length) errorText = unsupportedMessage(unsupported);
    else if (readFail) errorText = partialReadText(readFail);
    return {
      files: files.slice(),
      fileText: fileText,
      errorText: errorText,
      canStart: files.length > 0
    };
  }

  function failedSelection() {
    return { files: [], unsupported: [], problems: [{ name: '', reason: 'read' }] };
  }

  function createSelectionGate() {
    var current = 0;
    return {
      start: function () {
        current += 1;
        return current;
      },
      isCurrent: function (token) {
        return token === current;
      }
    };
  }

  function optionNumber(value, fallback) {
    var number = Number(value);
    if (!Number.isFinite(number) || number <= 0) return fallback;
    return number;
  }

  function readOptions(options) {
    return {
      timeoutMs: optionNumber(options && options.timeoutMs, DEFAULT_TIMEOUT_MS),
      maxBatches: Math.floor(optionNumber(options && options.maxBatches, DEFAULT_MAX_BATCHES))
    };
  }

  function withTimeout(promise, ms) {
    return new Promise(function (resolve, reject) {
      var settled = false;
      var timer = setTimeout(function () {
        if (settled) return;
        settled = true;
        var err = new Error('timeout');
        err.code = 'read';
        reject(err);
      }, ms);
      promise.then(function (value) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        resolve(value);
      }, function (err) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(err);
      });
    });
  }

  function readOneBatch(reader) {
    return new Promise(function (resolve, reject) {
      var settled = false;
      function finish(ok, value) {
        if (settled) return;
        settled = true;
        if (ok) resolve(Array.isArray(value) ? value.slice() : []);
        else reject(value instanceof Error ? value : new Error('readEntries failed'));
      }
      try {
        var returned = reader.readEntries(function (batch) {
          finish(true, batch);
        }, function (err) {
          finish(false, err || new Error('readEntries failed'));
        });
        if (returned && typeof returned.then === 'function') {
          returned.then(function (batch) {
            finish(true, batch);
          }, function (err) {
            finish(false, err);
          });
        }
      } catch (err) {
        finish(false, err);
      }
    });
  }

  async function readAllBatches(reader, options) {
    var all = [];
    for (var n = 0; n < options.maxBatches; n++) {
      var batch = await withTimeout(readOneBatch(reader), options.timeoutMs);
      if (!batch.length) return all;
      for (var i = 0; i < batch.length; i++) all.push(batch[i]);
    }
    var err = new Error('too many directory batches');
    err.code = 'read';
    err.partial = all;
    throw err;
  }

  function fileFromEntry(entry) {
    return new Promise(function (resolve, reject) {
      var settled = false;
      function finish(ok, value) {
        if (settled) return;
        settled = true;
        if (ok) resolve(value);
        else reject(value instanceof Error ? value : new Error('file failed'));
      }
      try {
        var returned = entry.file(function (file) {
          finish(true, file);
        }, function (err) {
          finish(false, err || new Error('file failed'));
        });
        if (returned && typeof returned.then === 'function') {
          returned.then(function (file) {
            finish(true, file);
          }, function (err) {
            finish(false, err);
          });
        }
      } catch (err) {
        finish(false, err);
      }
    });
  }

  async function walkEntry(entry, out, problems, options) {
    if (!entry) return;
    if (entry.isFile) {
      try {
        var file = await withTimeout(fileFromEntry(entry), options.timeoutMs);
        if (file) out.push(file);
      } catch (err) {
        problems.push({ name: entry.name || '파일', reason: 'read' });
      }
      return;
    }
    if (!entry.isDirectory) return;
    var partial = [];
    try {
      var reader = entry.createReader();
      if (!reader || typeof reader.readEntries !== 'function') throw new Error('no reader');
      var children = await readAllBatches(reader, options);
      partial = children;
    } catch (err) {
      if (err && err.partial) partial = err.partial;
      problems.push({ name: entry.name || '폴더', reason: 'read' });
    }
    for (var i = 0; i < partial.length; i++) await walkEntry(partial[i], out, problems, options);
  }

  function snapshotDataTransfer(dt) {
    var snapshot = { entries: [], looseFiles: [], entryMode: false };
    if (!dt) return snapshot;
    try {
      var files = dt.files;
      if (files && typeof files.length === 'number') {
        for (var i = 0; i < files.length; i++) snapshot.looseFiles.push(files[i]);
      }
    } catch (err) {
      snapshot.looseFiles = [];
    }
    var items = null;
    try {
      items = dt.items;
    } catch (err2) {
      items = null;
    }
    if (!items || typeof items.length !== 'number') return snapshot;
    for (var n = 0; n < items.length; n++) {
      var item = items[n];
      if (!item) continue;
      if (item.kind && item.kind !== 'file') {
        snapshot.entries.push({ skip: true, kind: item.kind });
        continue;
      }
      var entry = null;
      var file = null;
      var entryError = null;
      try {
        if (typeof item.webkitGetAsEntry === 'function') {
          entry = item.webkitGetAsEntry();
          snapshot.entryMode = true;
        } else if (typeof item.getAsEntry === 'function') {
          entry = item.getAsEntry();
          snapshot.entryMode = true;
        }
      } catch (err3) {
        entryError = err3;
        snapshot.entryMode = true;
      }
      if (!entry && typeof item.getAsFile === 'function') {
        try {
          file = item.getAsFile();
        } catch (err4) {
          entryError = entryError || err4;
        }
      }
      snapshot.entries.push({
        kind: item.kind || 'file',
        entry: entry,
        file: file,
        entryError: entryError
      });
    }
    return snapshot;
  }

  async function collectSnapshot(snapshot, options) {
    var opts = readOptions(options);
    var source = snapshot || { entries: [], looseFiles: [], entryMode: false };
    var raw = [];
    var problems = [];
    if (source.entryMode) {
      var sawEntry = false;
      var entries = source.entries || [];
      for (var i = 0; i < entries.length; i++) {
        var item = entries[i];
        if (!item || item.skip) continue;
        if (item.entry) {
          sawEntry = true;
          await walkEntry(item.entry, raw, problems, opts);
        } else if (item.file) raw.push(item.file);
        else if (item.entryError) problems.push({ name: '항목', reason: 'read' });
      }
      if (!sawEntry && !raw.length && source.looseFiles && source.looseFiles.length) {
        for (var j = 0; j < source.looseFiles.length; j++) raw.push(source.looseFiles[j]);
        problems.length = 0;
      }
    } else {
      var loose = source.looseFiles || [];
      for (var k = 0; k < loose.length; k++) raw.push(loose[k]);
    }
    return classifyFileList(raw, problems);
  }

  function detailMessage(detail) {
    var text = '';
    if (typeof detail === 'string') text = detail;
    else if (Array.isArray(detail)) {
      text = detail.map(function (item) {
        if (!item) return '';
        if (typeof item === 'string') return item;
        return item.msg || item.message || '';
      }).filter(Boolean).join(' ');
    }
    text = text.replace(/\s+/g, ' ').trim();
    return text ? text.slice(0, 300) : '';
  }

  async function submitJob(files, fetchImpl) {
    var list = toArray(files).filter(Boolean);
    if (!list.length) {
      var emptyErr = new Error(MESSAGES.uploadEmpty);
      emptyErr.code = 'empty';
      throw emptyErr;
    }
    if (typeof fetchImpl !== 'function') {
      var fetchErr = new Error(MESSAGES.uploadFailed);
      fetchErr.code = 'fetch';
      throw fetchErr;
    }
    var body = new FormData();
    for (var i = 0; i < list.length; i++) {
      body.append('files', list[i], list[i].name || '영수증');
    }
    var response;
    try {
      response = await fetchImpl('/api/jobs', {
        method: 'POST',
        body: body,
        credentials: 'same-origin'
      });
    } catch (cause) {
      var netErr = new Error(MESSAGES.uploadNetwork);
      netErr.code = 'network';
      netErr.cause = cause;
      throw netErr;
    }
    var payload = null;
    try {
      payload = await response.json();
    } catch (parseErr) {
      payload = null;
    }
    if (!response || !response.ok) {
      var httpErr = new Error(detailMessage(payload && payload.detail) || MESSAGES.uploadFailed);
      httpErr.code = 'http';
      httpErr.status = response && response.status;
      throw httpErr;
    }
    if (!payload || typeof payload.job_id !== 'string' || !payload.job_id) {
      var badErr = new Error(MESSAGES.uploadFailed);
      badErr.code = 'bad-response';
      throw badErr;
    }
    return payload;
  }

  return {
    ALLOWED_EXT: ALLOWED_EXT,
    MESSAGES: MESSAGES,
    extensionOf: extensionOf,
    isSupportedName: isSupportedName,
    classifyFileList: classifyFileList,
    describeSelection: describeSelection,
    failedSelection: failedSelection,
    createSelectionGate: createSelectionGate,
    snapshotDataTransfer: snapshotDataTransfer,
    collectSnapshot: collectSnapshot,
    submitJob: submitJob,
    unsupportedMessage: unsupportedMessage,
    selectedText: selectedText,
    partialReadText: partialReadText
  };
});
