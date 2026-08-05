// Flagged Summary 리포트의 Unflag 버튼을 처리하는 작고 독립적인 도우미 서버.
// flagged.js가 리포트를 만든 뒤 이 파일을 detached로 띄운다 → 감시자(watcher)가 안 켜져 있어도 Unflag가 동작.
// PC를 끄거나 로그아웃할 때까지 계속 실행된다. 같은 포트가 이미 열려 있으면(=이미 실행 중) 조용히 종료.
// 인자: <port> <clientId> <tenant> <secret?> <mailbox?>
const http = require('http');
const { unflagLocal, openLocal } = require('./outlook-read');

const PORT = parseInt(process.argv[2] || '3940', 10);
const CLIENT_ID = process.argv[3] || '';
const TENANT = process.argv[4] || '';
const SECRET = process.argv[5] || '';

function resetIdle(srv) { /* 자동 종료 없음 — PC를 끌 때까지 유지 */ }
function page(msg) {
  return `<!doctype html><meta charset="utf-8"><body style="font-family:'Segoe UI','Malgun Gothic',sans-serif;text-align:center;padding-top:60px;color:#333">`
    + `<div style="font-size:15pt">${msg}</div><script>setTimeout(function(){window.close()},1500)</script>`;
}

const srv = http.createServer(async (req, res) => {
  const u = new URL(req.url, `http://127.0.0.1:${PORT}`);
  res.setHeader('content-type', 'text/html; charset=utf-8');
  res.setHeader('Access-Control-Allow-Origin', '*'); // file:// 리포트에서 fetch 허용
  resetIdle(srv);
  try {
    if (u.pathname === '/ping') { res.end('ok'); return; }
    if (SECRET && u.searchParams.get('k') !== SECRET) { res.statusCode = 403; res.end(page('bad request')); return; }
    if (u.pathname === '/unflag') {
      const id = u.searchParams.get('id');
      if (!id) { res.statusCode = 400; res.end(page('no id')); return; }
      const r = await unflagLocal(id);
      res.end(page(r === 'gone' ? 'already gone' : '✅ Unflagged'));
      return;
    }
    if (u.pathname === '/open') {
      const id = u.searchParams.get('id');
      if (!id) { res.statusCode = 400; res.end(page('no id')); return; }
      await openLocal(id);
      res.end(page('Outlook 에서 열었습니다'));
      return;
    }
    res.statusCode = 404; res.end(page('?'));
  } catch (e) {
    const gone = /찾지 못|not found/i.test(e.message);
    res.statusCode = gone ? 200 : 500;
    res.end(page(gone ? 'already gone' : 'error: ' + String(e.message)));
  }
});
// 포트가 이미 쓰이면(=도우미가 이미 떠 있음) 조용히 종료 — 기존 것이 계속 처리
srv.on('error', () => { process.exit(0); });
srv.listen(PORT, '127.0.0.1', () => resetIdle(srv));
