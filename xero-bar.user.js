// ==UserScript==
// @name         XERO bar
// @namespace    xero-tools
// @version      2026.08.05.0500
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
// Result: the bar shows instantly. And when newer code arrives, it is swapped in RIGHT AWAY
// (no more "refresh twice" - see the hot-swap note in onload below).
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

  // 도구 코드가 로더의 나이를 확인할 수 있도록 버전을 남긴다.
  // 로더가 옛날이면 도구 코드가 바에 '업데이트 필요' 버튼을 띄운다 (조용히 옛 코드가 도는 일 방지).
  var LOADER_VER = '2026.08.05.0500';
  try { window.__xbarLoaderVer = LOADER_VER; } catch (e) {}

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
          var changed = false;
          if (hasGM) {
            var prev = '';
            try { prev = GM_getValue(CACHE_KEY, ''); } catch (e) {}
            if (body !== prev) { changed = true; try { GM_setValue(CACHE_KEY, body); } catch (e) {} }
          }
          if (!ran) { ran = true; runCode(body, CODE_URLS[i]); return; }   // no cache existed -> run fresh now

          // HOT-SWAP: newer code just arrived. Apply it now instead of waiting for the next page load.
          // Safe because the tool code: (1) returns early if #xbar already exists, so we remove the old
          // panel first; (2) registers the F13-F24 key listener only once, guarded by window.__xbarHK,
          // so no double-firing; (3) that listener dispatches through window.__xbarTOOLS, which the
          // re-run replaces -> the F-keys start using the new code immediately too.
          if (changed) {
            try { var oldbar = document.getElementById('xbar'); if (oldbar) oldbar.remove(); } catch (e) {}
            runCode(body, 'hot-swap');
          }
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
