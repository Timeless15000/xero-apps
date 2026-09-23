// Flagged Summary / Review Daily 리포트의 Unflag · 열기 버튼을 처리하는 작고 독립적인 도우미 서버.
// flagged.js / review-daily.js 가 리포트를 만든 뒤 이 파일을 detached로 띄운다 → 감시자(watcher)가 안 켜져 있어도 동작.
// PC를 끄거나 로그아웃할 때까지 계속 실행된다. 같은 포트가 이미 열려 있으면(=이미 실행 중) 조용히 종료.
//   /unflag?id=<Graph id>&mb=<사서함>          → Graph 로 flag 해제 (Outlook 종류·실행 여부 무관)
//   /open?id=<Graph id>&mb=<사서함>&web=<링크> → 클래식 Outlook 이 켜져 있으면 그 창에서 열고, 아니면 웹 Outlook 으로
// 인자: <port> <clientId> <tenant> <secret?>
const http = require('http');
const { setTenant } = require('./auth');
const graphRead = require('./graph-read');
const { outlookState } = require('./outlook-detect');

const PORT = parseInt(process.argv[2] || '3941', 10);
const CLIENT_ID = process.argv[3] || '';
const TENANT = process.argv[4] || '';
const SECRET = process.argv[5] || '';
if (TENANT) setTenant(TENANT);

function page(msg) {
  return `<!doctype html><meta charset="utf-8"><body style="font-family:'Segoe UI','Malgun Gothic',sans-serif;text-align:center;padding-top:60px;color:#333">`
    + `<div style="font-size:15pt">${msg}</div><script>setTimeout(function(){window.close()},1500)</script>`;
}

// 클래식 Outlook 창에서 열기 — Graph id 를 EntryID 로 바꿔 COM 으로 연다. 실패하면 false.
async function openInClassic(id, mailbox) {
  let st = 'none';
  try { st = await outlookState(); } catch (e) { st = 'none'; }
  if (st !== 'classic') return false;
  try {
    const entryId = await graphRead.toEntryId(CLIENT_ID, id, mailbox);
    if (!entryId) return false;
    const { openLocal } = require('./outlook-read');
    await openLocal(entryId);
    return true;
  } catch (e) { return false; }
}

const srv = http.createServer(async (req, res) => {
  const u = new URL(req.url, `http://127.0.0.1:${PORT}`);
  res.setHeader('content-type', 'text/html; charset=utf-8');
  res.setHeader('Access-Control-Allow-Origin', '*'); // file:// 리포트에서 fetch 허용
  try {
    if (u.pathname === '/ping') { res.end('ok'); return; }
    if (SECRET && u.searchParams.get('k') !== SECRET) { res.statusCode = 403; res.end(page('bad request')); return; }
    const id = u.searchParams.get('id');
    const mb = u.searchParams.get('mb') || '';
    if (u.pathname === '/unflag') {
      if (!id) { res.statusCode = 400; res.end(page('no id')); return; }
      await graphRead.unflag(CLIENT_ID, id, mb);
      res.end(page('✅ Unflagged'));
      return;
    }
    if (u.pathname === '/open') {
      if (!id) { res.statusCode = 400; res.end(page('no id')); return; }
      if (await openInClassic(id, mb)) { res.end(page('Outlook 에서 열었습니다')); return; }
      const web = u.searchParams.get('web') || '';
      if (/^https:\/\/outlook\.(office|office365|live|cloud\.microsoft)/i.test(web)) {
        res.statusCode = 302; res.setHeader('location', web); res.end(); return;
      }
      res.statusCode = 404; res.end(page('이 메일을 열 링크가 없습니다. 리포트를 다시 만들어 주세요.'));
      return;
    }
    res.statusCode = 404; res.end(page('?'));
  } catch (e) {
    const gone = /Graph 404|찾지 못|not found/i.test(e.message);
    res.statusCode = gone ? 200 : 500;
    res.end(page(gone ? 'already gone' : 'error: ' + String(e.message)));
  }
});
// 포트가 이미 쓰이면(=도우미가 이미 떠 있음) 조용히 종료 — 기존 것이 계속 처리
srv.on('error', () => { process.exit(0); });
srv.listen(PORT, '127.0.0.1');
