const assert = require('assert');
const path = require('path');
const { pathToFileURL } = require('url');
const { webPreferences, protectWindow, assertTrustedSender } = require('../../src/main/app-security');

const prefs = webPreferences();
assert.strictEqual(prefs.nodeIntegration, false);
assert.strictEqual(prefs.contextIsolation, true);
assert.strictEqual(prefs.preload, path.join(__dirname, '..', '..', 'src', 'main', 'preload.js'));
const frame = { url: pathToFileURL(path.join(__dirname, '..', '..', 'src', 'renderer', 'index.html')).href };
const contents = { mainFrame: frame };
const win = { webContents: contents };
assert.doesNotThrow(() => assertTrustedSender({ sender: contents, senderFrame: frame }, win));
assert.throws(() => assertTrustedSender({ sender: {}, senderFrame: frame }, win), /Untrusted/);
assert.throws(() => assertTrustedSender({ sender: contents, senderFrame: { url: frame.url } }, win), /Untrusted/);
assert.throws(() => assertTrustedSender({ sender: contents, senderFrame: null }, win), /Untrusted/);
assert.throws(() => assertTrustedSender({ sender: contents, senderFrame: frame }, null), /Untrusted/);
frame.url = 'https://example.com/';
assert.throws(() => assertTrustedSender({ sender: contents, senderFrame: frame }, win), /Untrusted/);

const handlers = {};
contents.setWindowOpenHandler = (handler) => { handlers.open = handler; };
contents.on = (name, handler) => { handlers[name] = handler; };
protectWindow(win);
assert.deepStrictEqual(handlers.open(), { action: 'deny' });
for (const name of ['will-navigate', 'will-frame-navigate', 'will-attach-webview']) {
  let blocked = false;
  handlers[name]({ preventDefault: () => { blocked = true; } });
  assert(blocked, `${name} must be blocked`);
}
console.log('Electron 보안 경계 단위 테스트 통과');
