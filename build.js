// Auto-generates xero-bar.user.js from Xero_applications.html
// Runs on GitHub Actions whenever Xero_applications.html is pushed.
// It opens the HTML in headless Chrome (same as opening it yourself),
// lets the page build the userscript, and saves the result.

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

    fs.writeFileSync(path.resolve(__dirname, 'xero-bar.user.js'), us, 'utf8');
    console.log('OK: wrote xero-bar.user.js (' + us.length + ' chars).');
  } finally {
    await browser.close();
  }
})().catch(e => { console.error(e); process.exit(1); });
