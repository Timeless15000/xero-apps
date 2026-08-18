// Review Daily — 최근 24시간 Inbox 메일 중 "답장 안 한" 메일을 HTML 리포트로 정리
// 실행: node src/index.js --review-daily  (OUTLOOK_bar.ahk의 Review Daily 버튼)
// 답장 판정: ①메일의 답장 표시(verb 102/103) ②같은 기간 '보낸 메일'에 같은 대화(ConversationID)가 있으면 답장한 것
// AI 호출 없음 — 빠르고 무료. noreply 등 자동 발송, 내가 보낸 메일, 전달-알림은 목록에서 뺀다.
const fs = require('fs');
const path = require('path');
const { execFile, spawn } = require('child_process');
const { readRecent } = require('./outlook-read');
const { getActiveOutlookAccount, outlookState } = require('./outlook-detect');

const ROOT = path.join(__dirname, '..');
const REPORT_DIR = path.join(ROOT, 'reports');
const LOG_FILE = path.join(ROOT, 'log.txt');

function log(msg) {
  const line = `[${new Date().toLocaleString('sv-SE')}] ${msg}`;
  console.log(line);
  try { fs.appendFileSync(LOG_FILE, line + '\n'); } catch (e) { }
}
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function slug(email) { return String(email || '').split('@')[0].replace(/[^A-Za-z0-9_-]/g, '') || 'shared'; }

const BUILTIN_SKIP = ['noreply', 'no-reply', 'no_reply', 'donotreply', 'do-not-reply', 'mailer-daemon', 'postmaster', 'notification', 'notifications@', 'alert@', 'alerts@'];

async function run(cfg, mailbox) {
  let target = String(mailbox || '').trim().toLowerCase();
  if (!target) {
    try { target = (await getActiveOutlookAccount()) || ''; } catch (e) { target = ''; }
  }
  const hours = cfg.reviewHours || 24;
  log(`📬 Review Daily 시작 — 사서함: ${target || '(기본)'}, 최근 ${hours}시간`);

  let rr;
  try {
    rr = await readRecent(target, { hours, bodyChars: 700, max: 400 });
  } catch (e) {
    let st = 'none';
    try { st = await outlookState(); } catch (e2) { }
    if (st === 'new') {
      throw new Error('이 PC는 "새 Outlook(New Outlook)"을 쓰고 있어 메일을 읽을 수 없습니다.\n'
        + 'Outlook 오른쪽 위의 "새 Outlook" 스위치를 꺼서 기존 Outlook으로 바꾼 뒤 다시 눌러주세요.');
    }
    if (st === 'none') {
      throw new Error('Outlook 이 실행되고 있지 않습니다. Outlook 을 켜고 다시 눌러주세요.');
    }
    throw new Error('Outlook 은 켜져 있는데 연결하지 못했습니다.\n원인: ' + e.message);
  }

  const boxArg = rr.mailbox || target || '';
  const skipList = [...new Set([...BUILTIN_SKIP, ...(cfg.skipSenders || [])])];
  // 내 주소 전부 — 사서함 주소 + Outlook 계정 주소들 (내가 보낸 메일이 Inbox로 들어오는 경우 제외용)
  const myAddrs = new Set([boxArg, ...(rr.myAddrs || [])]
    .map(a => String(a || '').toLowerCase()).filter(a => a.includes('@')));
  const total = rr.items.length;
  let replied = 0, auto = 0, mine = 0;
  const open = [];
  for (const m of rr.items) {
    const addr = String(m.addr || '').toLowerCase();
    if (m.mine || (addr && myAddrs.has(addr))) { mine++; continue; }           // 내가 보낸 메일
    if (skipList.some(s => addr.includes(s))) { auto++; continue; }            // 자동 발송
    if (m.verb === 102 || m.verb === 103) { replied++; continue; }             // 답장 표시 있음
    if (m.conv && rr.sentSet.has(m.conv)) { replied++; continue; }             // 보낸 메일에 같은 대화 있음
    open.push(m);
  }
  // 오래된 것부터
  open.sort((a, b) => new Date(a.received) - new Date(b.received));
  const received = Math.max(0, total - mine);   // 내가 보낸 메일은 '받은 메일' 수에서도 뺀다
  log(`📬 Review Daily — 받은 ${received}통 중 미답장 ${open.length}통 (답장됨 ${replied}, 자동발송 제외 ${auto}, 내 메일 제외 ${mine})`);

  // 제목 클릭 → Outlook에서 열기 (Flagged Summary와 같은 도우미 사용)
  const _uport = cfg.unflagPort || 3940;
  const _usecret = cfg.unflagSecret || '';
  const _ubase = `http://127.0.0.1:${_uport}`;
  try {
    const child = spawn(process.execPath,
      [path.join(__dirname, 'unflag-server.js'), String(_uport), cfg.clientId || '', cfg.tenant || '', _usecret],
      { detached: true, stdio: 'ignore', windowsHide: true });
    child.unref();
  } catch (e) { log(`⚠️  도우미 시작 실패: ${e.message}`); }
  for (const m of open) {
    m.webLink = `${_ubase}/open?id=${encodeURIComponent(m.id)}` + (_usecret ? `&k=${encodeURIComponent(_usecret)}` : '');
  }

  fs.mkdirSync(REPORT_DIR, { recursive: true });
  const d = new Date(), z = n => String(n).padStart(2, '0');
  const stamp = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}_${z(d.getHours())}${z(d.getMinutes())}`;
  const file = path.join(REPORT_DIR, `Review_Daily_${boxArg ? slug(boxArg) + '_' : ''}${stamp}.html`);
  fs.writeFileSync(file, render(open, { boxArg, hours, total: received, replied, auto, mine }));
  log(`📬 Review Daily 완료 — 리포트: ${path.basename(file)}`);
  execFile('cmd.exe', ['/c', 'start', '', file], { windowsHide: true });
}

function ageH(iso) { return Math.max(0, Math.round((Date.now() - new Date(iso).getTime()) / 3600000)); }

function render(open, info) {
  const from = new Date(Date.now() - info.hours * 3600000);
  const f2 = t => new Date(t).toLocaleString('en-AU', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
  let rows = '';
  for (const m of open) {
    const h = ageH(m.received);
    const cls = h >= 20 ? 'red' : h >= 8 ? 'or' : 'gray';
    const subject = esc(m.subject || '(no subject)');
    const title = m.webLink ? `<a href="${esc(m.webLink)}" target="_blank">${subject}</a>` : subject;
    const prev = esc(String(m.preview || '').replace(/\s+/g, ' ').slice(0, 220));
    rows += `<div class="mail">`
      + `<div class="top"><span class="badge ${cls}">${h}h</span>${m.unread ? '<span class="dot" title="unread"></span>' : ''}<span class="subj">${title}</span></div>`
      + `<div class="meta">${esc(m.name || m.addr || '')} &lt;${esc(m.addr || '')}&gt; · ${f2(m.received)}</div>`
      + `<div class="prev">${prev}</div>`
      + `</div>`;
  }
  if (!open.length) rows = `<div class="none">🎉 All replied — no unanswered emails in the last ${info.hours}h.<br>최근 ${info.hours}시간 안에 답장 안 한 메일이 없습니다.</div>`;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Review Daily</title><style>
body{font:14px/1.5 "Segoe UI","Malgun Gothic",Arial,sans-serif;color:#222;margin:0;background:#f2f4f5}
.hd{background:#C8511B;color:#fff;padding:16px 24px}
.hd .t{font-size:20px;font-weight:bold}
.hd small{display:block;font-size:12px;opacity:.95;margin-top:3px}
.wrap{max-width:980px;margin:0 auto;padding:18px 24px}
.sum-cards{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 16px}
.card{background:#fff;border:1px solid #d9dee3;border-radius:4px;padding:10px 16px;min-width:110px}
.card b{display:block;font-size:20px}
.card span{color:#666;font-size:12px}
.card.red b{color:#c1272d}
.mail{background:#fff;border:1px solid #d9dee3;border-radius:6px;padding:12px 16px;margin:0 0 10px}
.top{display:flex;align-items:center;gap:8px}
.subj{font-weight:bold;font-size:15px}
.subj a{color:#0F6CBD;text-decoration:none}
.subj a:hover{text-decoration:underline}
.badge{font-weight:bold;font-size:11px;padding:2px 8px;border-radius:10px;color:#fff;white-space:nowrap}
.badge.red{background:#c1272d}.badge.or{background:#e65100}.badge.gray{background:#8a949e}
.dot{width:9px;height:9px;border-radius:50%;background:#0F6CBD;display:inline-block}
.meta{color:#666;font-size:12px;margin:2px 0 6px}
.prev{color:#555;font-size:12.5px;border-left:3px solid #eadfd7;padding-left:10px}
.none{background:#fff;border:1px solid #d9dee3;border-radius:6px;padding:24px;text-align:center;font-size:16px}
.ft{color:#888;font-size:11px;margin:22px 0}
</style></head><body>
<div class="hd"><div class="t">📬 Review Daily — unanswered emails</div>
<small>${esc(info.boxArg || '(default mailbox)')} · ${f2(from)} → ${f2(new Date())} (last ${info.hours}h) · Outlook에서 직접 읽음</small></div>
<div class="wrap">
<div class="sum-cards">
<div class="card${open.length ? ' red' : ''}"><b>${open.length}</b><span>Unanswered / 미답장</span></div>
<div class="card"><b>${info.replied}</b><span>Replied / 답장됨</span></div>
<div class="card"><b>${info.total}</b><span>Received / 받은 메일</span></div>
<div class="card"><b>${info.auto}</b><span>Auto mail skipped</span></div>
<div class="card"><b>${info.mine || 0}</b><span>My own mail skipped</span></div>
</div>
${rows}
<p class="ft">Click a subject to open that email in Outlook. Replied = Outlook reply mark or a sent email in the same conversation within the window. Auto/no-reply senders and mail you sent yourself are excluded. 제목을 누르면 Outlook에서 열립니다.</p>
</div>
</body></html>`;
}

module.exports = { run, render };
