// 실시간 답장 초안 감시자 — 새 메일 도착 → 분류 → 톤 가이드 기반 초안 → Outlook 임시보관함 저장
const fs = require('fs');
const path = require('path');
const http = require('http');
const { login, setTenant } = require('./auth');
const { getMe, getNewMessages, getMessageText, createReplyDraft, prependSummary, removeSummary, findSummarizedMessages, unflagMessage } = require('./graph');
const { classify, draftReply, summarize } = require('./ai');

const ROOT = path.join(__dirname, '..');
const STATE_FILE = path.join(ROOT, 'state.json');
const LOG_FILE = path.join(ROOT, 'log.txt');

function loadJson(f, fallback) { try { return JSON.parse(fs.readFileSync(f, 'utf8')); } catch (e) { return fallback; } }
function log(msg) {
  const line = `[${new Date().toLocaleString('sv-SE')}] ${msg}`;
  console.log(line);
  try { fs.appendFileSync(LOG_FILE, line + '\n'); } catch (e) { }
}

const cfg = loadJson(path.join(ROOT, 'config.json'), null);
if (!cfg || !cfg.clientId || cfg.clientId.includes('여기에')) {
  console.error('⚠️  config.json이 없거나 clientId가 비어 있습니다.');
  console.error('   config.example.json을 config.json으로 복사한 뒤 값을 채워주세요. (README.md 참고)');
  process.exit(1);
}
setTenant(cfg.tenant);

const state = loadJson(STATE_FILE, { lastCheck: null, processed: [], stats: { drafted: 0, skipped: 0 } });
if (!Array.isArray(state.summarized)) state.summarized = []; // 한글 요약을 삽입한 메일 id 목록 (일괄 삭제용)
function saveState() { fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2)); }

function escHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

let myEmail = '';

// 설정과 무관하게 항상 걸러내는 자동 발송 주소 패턴 (config.json의 skipSenders와 합쳐서 적용)
const BUILTIN_SKIP = ['noreply', 'no-reply', 'no_reply', 'donotreply', 'do-not-reply', 'mailer-daemon', 'postmaster'];

async function processMessage(m) {
  const fromEmail = (m.from?.emailAddress?.address || '').toLowerCase();
  const fromName = m.from?.emailAddress?.name || fromEmail;
  const subject = m.subject || '(제목 없음)';

  if (!fromEmail || fromEmail === myEmail) return; // 내가 보낸 메일

  // CC로만 받은 메일은 참고용 — 초안 만들지 않음
  const addrOf = r => (r.emailAddress?.address || '').toLowerCase();
  const inTo = (m.toRecipients || []).some(r => addrOf(r) === myEmail);
  const inCc = (m.ccRecipients || []).some(r => addrOf(r) === myEmail);
  if (inCc && !inTo) {
    state.stats.skipped++;
    log(`⏭️  CC 참고 메일이라 건너뜀: "${subject}" (${fromName})`);
    return;
  }

  // 발신 주소와 답장 주소(Reply-To) 모두 검사 — AGL처럼 발신은 정상 주소인데
  // 답장만 noreply로 가는 자동 발송 메일을 걸러내기 위함
  const replyToEmail = (m.replyTo?.[0]?.emailAddress?.address || '').toLowerCase();
  const skipList = [...new Set([...BUILTIN_SKIP, ...(cfg.skipSenders || [])])];
  const skipHit = skipList.find(s => fromEmail.includes(s) || replyToEmail.includes(s));
  if (skipHit) {
    state.stats.skipped++;
    const via = replyToEmail && !fromEmail.includes(skipHit) ? ` (답장주소 ${replyToEmail})` : '';
    log(`⏭️  발신자 필터로 건너뜀: ${fromEmail}${via} — "${subject}"`);
    return;
  }

  // 1) 답장 필요 여부 분류
  const cls = await classify(cfg, { fromName, fromEmail, replyToEmail, subject, preview: m.bodyPreview || '' });
  if (!cls.needsReply) {
    state.stats.skipped++;
    log(`⏭️  답장 불필요(${cls.category}): "${subject}" — ${cls.reason}`);
    return;
  }

  // 2) 본문 전체를 읽고 초안 생성 (긴 스레드 요약을 위해 넉넉히 읽는다)
  const full = await getMessageText(cfg.clientId, m.id, 15000);
  const myDomain = myEmail.split('@')[1] || '';
  const isInternal = !!myDomain && fromEmail.endsWith('@' + myDomain);
  const reply = await draftReply(cfg, toneGuide(), {
    fromName, fromEmail, subject, text: full.text,
  }, cls.category, { language: cls.language, isInternal });

  // 3) Outlook 임시보관함에 답장 초안 저장 (서명 자동 첨부)
  // 요약 삽입보다 반드시 먼저 — 초안의 인용문에 요약이 딸려 들어가지 않도록
  const sig = signature();
  const replyHtml = escHtml(reply).replace(/\n/g, '<br>') + sig.html;
  const draftId = await createReplyDraft(cfg.clientId, m.id, replyHtml, sig.images);
  state.stats.drafted++;
  log(`✍️  초안 저장 완료 [${cls.category}] "${subject}" (${fromName})`);

  // 4) 긴 메일이면 영어 요약을 원본 메일과 답장 초안 "양쪽" 맨 위에 삽입
  //    (초안을 열자마자 요약이 보이도록 — 보내기 시에는 매크로가 자동 삭제)
  const minChars = cfg.summaryMinChars ?? 1500;
  if (full.text.length >= minChars) {
    try {
      const sum = await summarize(cfg, { fromName, subject, text: full.text });
      if (await prependSummary(cfg.clientId, m.id, renderSummaryHtml(sum, m.id, draftId, 'original'), sum)) {
        state.summarized.push(m.id);
        log(`📌 요약 삽입(원본): "${subject}"`);
      }
      if (await prependSummary(cfg.clientId, draftId, renderSummaryHtml(sum, m.id, draftId, 'draft'), sum, { allowDraft: true })) {
        state.summarized.push(draftId);
        log(`📌 요약 삽입(초안): "${subject}"`);
      }
    } catch (e) {
      log(`⚠️  요약 삽입 실패(초안은 정상 저장됨) "${subject}": ${e.message}`);
    }
  }
}

// "- 불릿" / "ACTION: ..." 형식의 요약 텍스트를 노란 박스 HTML로 변환 (Outlook에서 원문과 확실히 구분되도록)
// [제목] 줄도 계속 처리한다 — 이전 버전으로 만들어진 요약과의 호환용
// ✕ 링크는 감시자에 내장된 로컬 삭제 서버(startCleanServer)를 호출한다
// variant: 'original' = 받은 메일용(✕ 링크 표시) / 'draft' = 답장 초안용
// (열려 있는 작성 창은 링크로 못 지우므로, 초안에는 링크 대신 상단 [요약 지우기] 버튼 안내를 넣는다)
function renderSummaryHtml(sumText, msgId, draftId, variant = 'original') {
  let body = '';
  for (const raw of sumText.split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    const title = line.match(/^\[(.+?)\]$/);
    if (title) body += `<div style="font-weight:bold;margin:10px 0 2px 0;color:#5B4A00;">${escHtml(title[1])}</div>`;
    else if (/^ACTION\s*:/i.test(line)) body += `<div style="font-weight:bold;margin:9px 0 0 0;color:#8A4A00;">${escHtml(line)}</div>`;
    else body += `<div style="margin:0 0 2px 0;">${escHtml(line)}</div>`;
  }
  const k = CLEAN_SECRET ? `&k=${encodeURIComponent(CLEAN_SECRET)}` : '';
  const rightSide = variant === 'draft'
    ? `<span style="color:#B0A264;font-size:9pt;">Remove: <b>✕ Clear summary</b> button at top of window · auto-removed on send</span>`
    : `<a href="${CLEAN_BASE}/remove?id=${encodeURIComponent(msgId)}${draftId ? `&draft=${encodeURIComponent(draftId)}` : ''}${k}" title="Clear the summary from this email and its reply draft" style="color:#8A6D00;text-decoration:none;font-weight:bold;font-size:12pt;">✕ Clear</a>`
    + `<span style="color:#D8C97A;font-size:9pt;">&nbsp;|&nbsp;</span>`
    + `<a href="${CLEAN_BASE}/remove-all?x=1${k}" title="Clear summaries from every email" style="color:#B0A264;text-decoration:none;font-size:9pt;">Clear all</a>`;
  const header = `<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>`
    + `<td style="font-weight:bold;font-size:11pt;color:#8A6D00;font-family:Calibri,Arial,sans-serif;">📌 AI Thread Summary <span style="font-weight:normal;font-size:9pt;color:#999;">(auto-generated · reference only)</span></td>`
    + `<td align="right" style="white-space:nowrap;font-family:Calibri,Arial,sans-serif;">`
    + rightSide
    + `</td></tr></table>`;
  return `<div style="background:#FFF9DB;border:1px solid #E0C200;border-radius:8px;padding:12px 16px;margin:0 0 18px 0;font-family:Calibri,Arial,sans-serif;font-size:10.5pt;line-height:1.5;color:#333;">`
    + header + body + `</div>`;
}

// ✕ 클릭을 받아 요약을 지워주는 미니 서버
// 기본은 로컬 전용(127.0.0.1). 클라우드 서버에서 돌릴 때는 config.json에
// cleanBaseUrl(외부 주소)과 cleanSecret(비밀 키)을 넣으면 외부에서도 받되 키로 보호한다
const CLEAN_PORT = cfg.cleanPort ?? 3939;
const CLEAN_BASE = (cfg.cleanBaseUrl || `http://127.0.0.1:${CLEAN_PORT}`).replace(/\/+$/, '');
const CLEAN_SECRET = cfg.cleanSecret || '';
function cleanPage(msg) {
  return `<!doctype html><meta charset="utf-8"><title>요약 삭제</title>`
    + `<body style="font-family:'Malgun Gothic',sans-serif;text-align:center;padding-top:70px;color:#333;">`
    + `<div style="font-size:16pt;">${msg}</div>`
    + `<div style="font-size:10pt;color:#999;margin-top:14px;">이 창은 곧 자동으로 닫힙니다. Outlook 화면에는 잠시 후 반영됩니다 — 안 지워져 보이면 다른 메일을 눌렀다가 다시 여세요.<br>이미 열려 있는 답장 작성 창에는 반영되지 않을 수 있지만, 보내기(Send)를 누르면 매크로가 요약을 자동 삭제한 뒤 발송하니 안심하세요.</div>`
    + `<script>setTimeout(function(){window.close()},2500)</script>`;
}
function startCleanServer() {
  const srv = http.createServer(async (req, res) => {
    const u = new URL(req.url, `http://127.0.0.1:${CLEAN_PORT}`);
    res.setHeader('content-type', 'text/html; charset=utf-8');
    res.setHeader('Access-Control-Allow-Origin', '*'); // Flagged Summary 리포트(file://)의 Unflag fetch 허용
    if (CLEAN_SECRET && u.searchParams.get('k') !== CLEAN_SECRET) {
      res.statusCode = 403;
      res.end(cleanPage('⚠️ 잘못된 요청입니다.'));
      return;
    }
    try {
      if (u.pathname === '/remove') {
        // 원본 메일과 답장 초안 양쪽에서 요약 제거 (한쪽이 이미 삭제/이동됐어도 나머지는 진행)
        const targets = [u.searchParams.get('id'), u.searchParams.get('draft')].filter(Boolean);
        let removed = 0, missing = 0;
        for (const t of targets) {
          try { if (await removeSummary(cfg.clientId, t) === 'removed') removed++; }
          catch (e) { if (/Graph 404/.test(e.message)) missing++; else throw e; }
        }
        log(`🖱️  ✕ 클릭 — 요약 삭제 ${removed}건${missing ? `, 메일 없음 ${missing}건` : ''}`);
        res.end(cleanPage(removed ? '✅ 요약을 지웠습니다. (원본·답장 초안 모두)' : 'ℹ️ 이미 지워진 요약입니다.'));
      } else if (u.pathname === '/remove-all') {
        await cleanSummaries();
        res.end(cleanPage('🧹 모든 메일의 요약을 지웠습니다.'));
      } else if (u.pathname === '/unflag') {
        // Flagged Summary 리포트의 Unflag 버튼 — 해당 메일의 flag 해제
        const id = u.searchParams.get('id');
        if (!id) { res.statusCode = 400; res.end(cleanPage('⚠️ id가 없습니다.')); return; }
        await unflagMessage(cfg.clientId, id);
        log(`🖱️  Unflag 클릭 — flag 해제 ${id.slice(0, 12)}…`);
        res.end(cleanPage('✅ Unflag 완료 — flag가 해제됐습니다.'));
      } else {
        res.statusCode = 404;
        res.end(cleanPage('알 수 없는 요청입니다.'));
      }
    } catch (e) {
      const gone = /Graph 404/.test(e.message);
      if (!gone) log(`⚠️  요약 삭제(✕) 실패: ${e.message}`);
      res.statusCode = gone ? 200 : 500;
      res.end(cleanPage(gone ? 'ℹ️ 메일이 이동/삭제되어 찾지 못했습니다. "전체 지우기"를 사용해보세요.' : `⚠️ 삭제 실패: ${escHtml(e.message)}`));
    }
  });
  srv.on('error', e => log(`⚠️  삭제 버튼 서버를 켜지 못했습니다(포트 ${CLEAN_PORT}): ${e.message} — ✕ 클릭이 작동하지 않을 수 있습니다`));
  srv.listen(CLEAN_PORT, cfg.cleanBaseUrl ? '0.0.0.0' : '127.0.0.1');
}

// 삽입해 둔 한글 요약을 전부 제거 (npm run clean)
async function cleanSummaries() {
  const ids = new Set(state.summarized);
  try {
    // 검색으로도 회수: state에서 빠졌거나 폴더 이동으로 id가 바뀐 메일까지
    for (const m of await findSummarizedMessages(cfg.clientId)) ids.add(m.id);
  } catch (e) { log(`⚠️  요약 메일 검색 실패(목록 기반으로만 진행): ${e.message}`); }

  let removed = 0, none = 0, gone = 0, failed = 0;
  for (const id of ids) {
    try {
      (await removeSummary(cfg.clientId, id)) === 'removed' ? removed++ : none++;
    } catch (e) {
      if (/Graph 404/.test(e.message)) gone++;
      else { failed++; log(`❌ 요약 제거 실패 (${id.slice(0, 12)}…): ${e.message}`); }
    }
  }
  if (failed === 0) state.summarized = [];
  saveState();
  log(`🧹 요약 일괄 삭제 완료 — 제거 ${removed}건, 이미 없음 ${none}건, 메일 삭제됨 ${gone}건${failed ? `, 실패 ${failed}건` : ''}`);
}

// 서명: signature.html + signature-logo.png (수정하면 다음 초안부터 반영)
let sigCache = null, sigMtime = 0;
function signature() {
  const f = path.join(ROOT, 'signature.html');
  try {
    const mt = fs.statSync(f).mtimeMs;
    if (!sigCache || mt !== sigMtime) {
      // Outlook이 서명 블록으로 인식하도록 표식을 붙인다
      // (_MailAutoSig 북마크: 데스크톱 Outlook 서명 교체 메뉴 / div id=Signature: 웹 Outlook)
      const html = `<div id="Signature"><a name="_MailAutoSig"></a><span style="mso-bookmark:_MailAutoSig">`
        + fs.readFileSync(f, 'utf8')
        + `</span></div>`;
      const images = [];
      const logo = path.join(ROOT, 'signature-logo.png');
      if (html.includes('cid:sig-logo') && fs.existsSync(logo)) {
        images.push({ name: 'signature-logo.png', contentId: 'sig-logo', contentBytes: fs.readFileSync(logo).toString('base64') });
      }
      sigCache = { html, images };
      sigMtime = mt;
    }
    return sigCache;
  } catch (e) { return { html: '', images: [] }; }
}

let toneCache = null, toneMtime = 0;
function toneGuide() {
  const f = path.join(ROOT, 'tone-guide.md');
  const mt = fs.statSync(f).mtimeMs;
  if (!toneCache || mt !== toneMtime) { toneCache = fs.readFileSync(f, 'utf8'); toneMtime = mt; }
  return toneCache;
}

async function poll() {
  const since = state.lastCheck;
  const nowIso = new Date().toISOString();
  const msgs = await getNewMessages(cfg.clientId, since);
  // 오래된 것부터 처리
  for (const m of msgs.reverse()) {
    if (m.isDraft || state.processed.includes(m.id)) continue;
    try {
      await processMessage(m);
    } catch (e) {
      log(`❌ 처리 실패 "${m.subject}": ${e.message}`);
    }
    state.processed.push(m.id);
    while (state.processed.length > 1000) state.processed.shift();
  }
  state.lastCheck = nowIso;
  saveState();
  // (초안 청소 루틴은 제거됨 — 이제 답장 초안에 요약을 의도적으로 넣고,
  //  보내기 직전 삭제는 클래식 Outlook 매크로 Application_ItemSend가 담당한다)
}

async function main() {
  const args = process.argv.slice(2);
  // 바의 확인창용 — Outlook 화면에서 감지한 사서함만 출력하고 끝낸다 (네트워크 호출 없음)
  if (args.includes('--which-mailbox')) {
    const { getActiveOutlookAccount } = require('./outlook-detect');
    let found = '';
    try { found = (await getActiveOutlookAccount()) || ''; } catch (e) { found = ''; }
    process.stdout.write('MAILBOX=' + found + '\n');
    return;
  }
  // 바의 사서함 선택 메뉴용 — Outlook에 붙어 있는 사서함 목록 + 지금 보고 있는 사서함
  if (args.includes('--list-mailboxes')) {
    const { getActiveOutlookAccount, listOutlookMailboxes, outlookState } = require('./outlook-detect');
    let list = [], cur = '';
    try { list = await listOutlookMailboxes(); } catch (e) { list = []; }
    if (!list.length) {
      let st = 'none';
      try { st = await outlookState(); } catch (e) { st = 'none'; }
      process.stdout.write('STATE=' + st + '\n');
    }
    try { cur = (await getActiveOutlookAccount()) || ''; } catch (e) { cur = ''; }
    // 지금 보고 있는 사서함을 맨 위로, 나머지는 이름(없으면 주소) 가나다/알파벳 순
    // (주소를 못 구한 사서함은 smtp가 비어 있고 이름만 있다 — cur도 이름일 수 있다)
    const key = b => (b.name || b.smtp || '').toLowerCase();
    list.sort((a, b) => key(a).localeCompare(key(b), 'en'));
    const curLow = String(cur || '').toLowerCase();
    const isCur = b => !!curLow && ((b.smtp && b.smtp === curLow) || (b.name && b.name.toLowerCase() === curLow));
    const top = list.filter(isCur);
    const rest = list.filter(b => !isCur(b));
    process.stdout.write('CUR=' + cur + '\n');
    for (const b of [...top, ...rest]) process.stdout.write('MBOX=' + (b.smtp || '') + '\t' + (b.name || '') + '\n');
    return;
  }
  // 바의 폴더 선택 창용 — 대상 사서함의 Inbox 하위 폴더 목록 (FLD=<spec>\t<표시 경로>)
  if (args.includes('--list-folders')) {
    const { listFolders } = require('./outlook-read');
    const _mi2 = args.indexOf('--mailbox');
    const _mb2 = _mi2 >= 0 ? (args[_mi2 + 1] || '') : (process.env.OBAR_MBOXPICK || '');
    try {
      for (const f of await listFolders(_mb2)) process.stdout.write('FLD=' + f.spec + '\t' + f.display + '\n');
    } catch (e) {
      process.stdout.write('ERR=' + (e.message || 'FAILED') + '\n');
    }
    return;
  }
  if (args.includes('--login')) { await login(cfg.clientId); return; }
  if (args.includes('--clean-summaries')) { await cleanSummaries(); return; }
  if (args.includes('--flagged-summary')) {
    // --mailbox <주소 또는 스토어 이름> 을 주면 그 사서함을 요약한다.
    // 한글 이름은 명령줄 인코딩이 깨질 수 있어 바가 환경변수 OBAR_MBOXPICK 으로도 전달한다.
    const _mi = args.indexOf('--mailbox');
    const _mbox = _mi >= 0 ? (args[_mi + 1] || '') : (process.env.OBAR_MBOXPICK || '');
    // --folders-file <경로>: 요약할 폴더 spec 목록 (UTF-8, 한 줄에 하나; '.'=Inbox) — 바의 폴더 선택 창이 만든다
    let _folders = [];
    const _fi = args.indexOf('--folders-file');
    if (_fi >= 0 && args[_fi + 1]) {
      try {
        _folders = fs.readFileSync(args[_fi + 1], 'utf8').replace(/^\uFEFF/, '')
          .split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      } catch (e) { log(`⚠️  폴더 목록 파일을 읽지 못했습니다(${e.message}) — Inbox만 요약합니다.`); }
    }
    await require('./flagged').run(cfg, _mbox, _folders);
    return;
  }
  // Review Daily — 최근 24시간 Inbox에서 답장 안 한 메일 리포트
  if (args.includes('--review-daily')) {
    const _mi4 = args.indexOf('--mailbox');
    const _mb4 = _mi4 >= 0 ? (args[_mi4 + 1] || '') : (process.env.OBAR_MBOXPICK || '');
    await require('./review-daily').run(cfg, _mb4);
    return;
  }

  const me = await getMe(cfg.clientId);
  myEmail = (me.mail || me.userPrincipalName || '').toLowerCase();
  log(`🚀 감시 시작 — 계정: ${myEmail}, 주기: ${cfg.pollSeconds || 30}초`);

  if (!state.lastCheck) {
    // 최초 실행: 과거 메일은 건드리지 않고 지금부터 감시
    state.lastCheck = new Date().toISOString();
    saveState();
    log('ℹ️  최초 실행 — 지금부터 도착하는 새 메일만 처리합니다.');
  }

  if (args.includes('--once')) { await poll(); log('1회 확인 완료.'); return; }

  startCleanServer(); // 요약 박스의 ✕ 링크 처리용
  while (true) {
    try { await poll(); } catch (e) {
      if (/fetch failed|network|timeout|abort|ETIMEDOUT|ECONN|ENOTFOUND/i.test(e.message)) {
        log('🌐 인터넷 연결 일시 불안정 — 다음 주기에 자동 재시도합니다 (메일 놓치지 않음)');
      } else {
        log('❌ 폴링 오류: ' + e.message);
      }
    }
    await new Promise(r => setTimeout(r, (cfg.pollSeconds || 30) * 1000));
  }
}

main().catch(e => { log('❌ 치명적 오류: ' + e.message); process.exit(1); });
