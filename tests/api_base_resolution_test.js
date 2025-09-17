const assert = require('assert');

function createStubElement() {
  return {
    style: {},
    addEventListener: () => {},
    removeEventListener: () => {},
    setAttribute: () => {},
    getAttribute: () => null,
    querySelector: () => null,
    querySelectorAll: () => [],
    classList: {
      toggle: () => {},
      add: () => {},
      remove: () => {}
    },
    appendChild: () => {},
    removeChild: () => {},
    textContent: '',
    disabled: false,
    ariaLabel: '',
    set ariaLabel(value) {
      this._ariaLabel = value;
    },
    get ariaLabel() {
      return this._ariaLabel;
    }
  };
}

function setupDom() {
  const elementCache = new Map();
  const getElement = (id) => {
    if (!elementCache.has(id)) {
      elementCache.set(id, createStubElement());
    }
    return elementCache.get(id);
  };

  global.window = {
    location: {
      protocol: 'https:',
      hostname: 'app.example.com',
      port: '',
      origin: 'https://app.example.com',
      search: ''
    },
    matchMedia: () => ({
      matches: false,
      addEventListener: () => {},
      removeEventListener: () => {},
      addListener: () => {},
      removeListener: () => {}
    })
  };

  global.document = {
    addEventListener: () => {},
    getElementById: getElement,
    querySelector: () => null,
    querySelectorAll: () => [],
    createElement: () => createStubElement(),
    documentElement: {
      setAttribute: () => {},
      removeAttribute: () => {}
    },
    body: {
      appendChild: () => {},
      removeChild: () => {}
    }
  };

  global.localStorage = {
    _store: new Map(),
    getItem(key) {
      return this._store.has(key) ? this._store.get(key) : null;
    },
    setItem(key, value) {
      this._store.set(key, String(value));
    },
    removeItem(key) {
      this._store.delete(key);
    }
  };
}

async function testUsesReachableCandidate() {
  setupDom();
  window.API_BASE_URL = 'https://api.example.com';

  const calls = [];
  global.fetch = async (url) => {
    calls.push(url);
    if (url.startsWith('https://api.example.com')) {
      throw new Error('unreachable');
    }
    if (url.startsWith('https://app.example.com')) {
      return await new Promise((resolve) => setTimeout(() => resolve({ ok: true }), 20));
    }
    throw new Error(`Unexpected URL ${url}`);
  };

  delete require.cache[require.resolve('../app.js')];
  const { ExcelProcessor } = require('../app.js');

  const processor = new ExcelProcessor();
  const resolved = await processor.ensureApiBaseUrl();

  assert.strictEqual(resolved, 'https://app.example.com');
  assert.strictEqual(processor.apiBaseUrl, 'https://app.example.com');
  assert.deepStrictEqual(calls, [
    'https://api.example.com/health',
    'https://app.example.com/health'
  ]);
}

async function testFallsBackWhenAllFail() {
  setupDom();
  delete window.API_BASE_URL;

  const calls = [];
  global.fetch = async (url) => {
    calls.push(url);
    throw new Error('network error');
  };

  delete require.cache[require.resolve('../app.js')];
  const { ExcelProcessor } = require('../app.js');

  const processor = new ExcelProcessor();
  const resolved = await processor.ensureApiBaseUrl();

  assert.strictEqual(resolved, 'https://app.example.com');
  assert.strictEqual(processor.apiBaseUrl, 'https://app.example.com');
  assert.ok(calls.length > 0, 'Expected fetch to be attempted for candidates');
}

async function run() {
  await testUsesReachableCandidate();
  await testFallsBackWhenAllFail();
  console.log('API base resolution tests passed.');
}

run().catch((error) => {
  console.error('API base resolution tests failed:', error);
  process.exitCode = 1;
});
