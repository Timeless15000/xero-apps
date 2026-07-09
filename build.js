// Auto-generates xero-bar.code.js (the tool payload) from Xero_applications.html
// Runs on GitHub Actions whenever Xero_applications.html is pushed.
// It opens the HTML in headless Chrome (same as opening it yourself),
// lets the page build the userscript, strips the header, and saves the code.
// The installed userscript (xero-bar.user.js) is a loader that fetches this file.

const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

(async () => {
  const htmlPath = path.resolve(__dirname, 'Xero_applications.html');
  if (!fs.existsSync(htmlPath)) {
    console.error('ERROR: Xero_applications.html not found in repo root.');
    process.exit(1);
  }

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  try {
    const page = await browser.newPage();
    await page.goto('file://' + htmlPath, { waitUntil: 'networkidle0' });
    // give the in-page generator a moment to fill the box
    await new Promise(r => setTimeout(r, 2000));

    const us = await page.evaluate(() => {
      const t = document.getElementById('usbox');
      return t ? t.value : '';
    });

    if (!us || us.indexOf('==UserScript==') < 0) {
      console.error('ERROR: userscript box was empty or invalid.');
      process.exit(1);
    }

    // Strip the ==UserScript== header. The loader (xero-bar.user.js) eval()s this
    // file, so it must be plain code and must NOT be installable on its own.
    const code = us.replace(/\/\/ ==UserScript==[\s\S]*?\/\/ ==\/UserScript==\s*/, '');
    fs.writeFileSync(path.resolve(__dirname, 'xero-bar.code.js'), code, 'utf8');
    console.log('OK: wrote xero-bar.code.js (' + code.length + ' chars).');
  } finally {
    await browser.close();
  }
})().catch(e => { console.error(e); process.exit(1); });
