/**
 * SermonCast NDI Sender — JS wrapper
 * Works in both development (npm start) and packaged Electron app.
 */

const path = require('path');
const fs   = require('fs');

let _addon     = null;
let _available = false;
let _error     = null;

function tryLoad() {
  const candidates = [
    // Development: ndi-addon/build/Release/ndi_sender.node
    path.join(__dirname, 'build', 'Release', 'ndi_sender.node'),
    // Packaged app: resources/ndi_sender.node (electron-builder extraResources)
    path.join(process.resourcesPath || '', 'ndi_sender.node'),
    // Same directory as this file
    path.join(__dirname, 'ndi_sender.node'),
  ];

  for (const p of candidates) {
    try {
      if (fs.existsSync(p)) {
        return require(p);
      }
    } catch (e) {
      _error = e.message;
    }
  }
  return null;
}

try {
  _addon     = tryLoad();
  _available = _addon !== null;
} catch (e) {
  _error = e.message;
}

module.exports = {
  isAvailable()              { return _available; },
  loadError()                { return _error; },
  createSender(n,w,h,fn,fd) { return _addon ? _addon.createSender(n,w,h,fn,fd) : false; },
  sendBGRA(buf)              { return _addon ? _addon.sendBGRA(buf) : false; },
  destroySender()            { if (_addon) _addon.destroySender(); },
  isReady()                  { return _addon ? _addon.isReady() : false; },
};
