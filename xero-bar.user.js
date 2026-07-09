// ==UserScript==
// @name         XERO bar
// @namespace    xero-tools
// @version      2026.07.09.1700
// @description  Always-latest loader for the XERO bar tools. Shows the bar INSTANTLY from a local cache, then refreshes the code in the background so the next page load has the newest version. Staff never reinstall or wait.
// @author       Timeless
// @match        https://go.xero.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      raw.githubusercontent.com
// @connect      cdn.jsdelivr.net
// @downloadURL  https://raw.githubusercontent.com/Timeless15000/xero-apps/main/xero-bar.user.js
// @updateURL    https://raw.githubusercontent.com/Timeless15000/xero-apps/main/xero-bar.user.js
// @run-at       document-idle
// ==/UserScript==
//
// HOW THIS WORKS (read me before editing):
// This file is a tiny LOADER. It does not contain the tools themselves.
// To make the bar appear with NO delay, it keeps the last-downloaded tool code in
// Tampermonkey storage and runs THAT immediately on every Xero page (stale-while-revalidate),
// then quietly downloads the newest xero-bar.code.js in the background and saves it for next time.
// Result: the bar shows instantly, and a code change appears on the *next* page load.
// (For an instant check of a change you just pushed, use the Increase Apply *bookmarklet* on the
//  Xero apps page - that always runs the newest code with no cache.)
// To change the tools, edit Xero_applications.html -> push -> the GitHub Action rebuilds
// xero-bar.code.js. NOBODY reinstalls. Only edit THIS loader if the loading mechanism changes
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
  var CACHE_KEY = 'xbar_code_cache';

  if (typeof GM_xmlhttpRequest === 'undefined') {
    console.error('[XERO bar] GM_xmlhttpRequest unavailable - please update Tampermonkey and re-install the loader.');
    return;
  }

  var hasGM = (typeof GM_getValue === 'function' && typeof GM_setValue === 'function');
  var ran = false;   // has the bar code been executed on this page yet?

  function valid(code) {
    return !!code && code.indexOf('xbar') !== -1;   // sanity check: real tool code contains "xbar"
  }

  function runCode(code, from) {
    try {
      (0, eval)(code);   // indirect eval -> runs in the userscript sandbox, bypasses the page CSP
    } catch (e) {
      console.error('[XERO bar] error running code from ' + from, e);
    }
  }

  // 1) INSTANT: run the cached code right away so the bar appears with no network wait.
  if (hasGM) {
    var cached = '';
    try { cached = GM_getValue(CACHE_KEY, ''); } catch (e) {}
    if (valid(cached)) { ran = true; runCode(cached, 'cache'); }
  }

  // 2) BACKGROUND: fetch the newest code, save it for next time. Only run it now if nothing ran yet
  //    (first-ever load / no cache) - otherwise the update applies on the next page load.
  function tryLoad(i) {
    if (i >= CODE_URLS.length) {
      if (!ran) console.error('[XERO bar] could not load tool code from any source.');
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
        if (res.status >= 200 && res.status < 300 && valid(body)) {
          if (hasGM) {
            var prev = '';
            try { prev = GM_getValue(CACHE_KEY, ''); } catch (e) {}
            if (body !== prev) { try { GM_setValue(CACHE_KEY, body); } catch (e) {} }
          }
          if (!ran) { ran = true; runCode(body, CODE_URLS[i]); }   // no cache existed -> run fresh now
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
