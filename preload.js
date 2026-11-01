// 앱 코드는 격리된 preload 세계에서만 Node와 IPC에 접근한다.
window.addEventListener('DOMContentLoaded', () => {
  const { pathToFileURL } = require('url');
  const path = require('path');
  if (window === window.top && window.location.href === pathToFileURL(path.join(__dirname, 'index.html')).href) {
    require('./renderer.js');
  }
});
