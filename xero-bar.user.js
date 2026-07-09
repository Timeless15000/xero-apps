// ==UserScript==
// @name         XERO bar
// @namespace    xero-tools
// @version      2026.07.09.1000
// @description  Always-latest loader for the XERO bar tools. Fetches the newest code from GitHub on every Xero page load, so staff never have to reinstall or wait for updates.
// @author       Timeless
// @match        https://go.xero.com/*
// @grant        GM_xmlhttpRequest
// @connect      raw.githubusercontent.com
// @connect      cdn.jsdelivr.net
// @downloadURL  https://raw.githubusercontent.com/Timeless15000/xero-apps/main/xero-bar.user.js
// @updateURL    https://raw.githubusercontent.com/Timeless15000/xero-apps/main/xero-bar.user.js
// @run-at       document-idle
// ==/UserScript==
//
// HOW THIS WORKS (read me before editing):
// This file is a tiny LOADER. It does not contain the tools themselves.
// On every Xero page it downloads the latest xero-bar.code.js from GitHub and runs it.
// So to change the tools, you edit Xero_applications.html -> push -> the GitHub Action
// rebuilds xero-bar.code.js -> every staff PC runs the new code on the next page load.
// NOBODY needs to reinstall. Only edit THIS loader if the loading mechanism itself changes
// (and if you do, bump @version above so installed copies pick up the new loader).
//
(function () {
  'use strict';
  if (window.__xbarLoaderRan) return;      // guard against double-injection
  window.__xbarLoaderRan = true;

  var CODE_URLS = [
    'https://raw.githubusercontent.com/Timeless15000/xero-apps/main/xero-bar.code.js',
    'https://cdn.jsdelivr.net/gh/Timeless15000/xero-apps@main/xero-bar.code.js'   // fallback if GitHub raw is down
  ];

  if (typeof GM_xmlhttpRequest === 'undefined') {
    console.error('[XERO bar] GM_xmlhttpRequest unavailable - please update Tampermonkey and re-install the loader.');
    return;
  }

  function runCode(code, from) {
    try {
      (0, eval)(code);   // indirect eval -> runs in the userscript sandbox, bypasses the page CSP
    } catch (e) {
      console.error('[XERO bar] error running code from ' + from, e);
    }
  }

  function tryLoad(i) {
    if (i >= CODE_URLS.length) {
      console.error('[XERO bar] could not load tool code from any source.');
      return;
    }
    var url = CODE_URLS[i] + '?_=' + Date.now();   // cache-buster -> always the newest code
    GM_xmlhttpRequest({
      method: 'GET',
      url: url,
      headers: { 'Cache-Control': 'no-cache', 'Pragma': 'no-cache' },
      timeout: 15000,
      onload: function (res) {
        var body = (res && res.responseText) ? res.responseText : '';
        if (res.status >= 200 && res.status < 300 && body.indexOf('xbar') !== -1) {
          runCode(body, CODE_URLS[i]);
        } else {
          tryLoad(i + 1);
        }
      },
      onerror: function () { tryLoad(i + 1); },
      ontimeout: function () { tryLoad(i + 1); }
    });
  }

  tryLoad(0);
})();
