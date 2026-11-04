// 테스트 전용: 다른 창에서 실제 IPC 핸들러에 접근을 시도한다.
const { ipcRenderer } = require('electron');

window.addEventListener('DOMContentLoaded', async () => {
  const calls = [
    ['dialog:openFile', {}],
    ['dialog:saveFile', {}],
    ['app:getUserDataPath'],
    ['shell:showItemInFolder', 'test'],
  ];
  window.__ipcResults = await Promise.all(calls.map(async ([channel, ...args]) => {
    try {
      await ipcRenderer.invoke(channel, ...args);
      return false;
    } catch (error) {
      return String(error).includes('Untrusted IPC sender');
    }
  }));
});
