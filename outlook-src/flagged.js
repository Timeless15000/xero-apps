// Flagged Email Summary — flag된 메일(마무리 안 된 일)을 AI 요약과 함께 HTML 리포트로 정리
// 실행: npm run flagged  (또는 OUTLOOK_bar.ahk의 Flagged Summary 버튼)
const fs = require('fs');
const path = require('path');
const { execFile, spawn } = require('child_process');
const { readFlagged } = require('./outlook-read');
const { getActiveOutlookAccount, outlookState } = require('./outlook-detect');
const { summarizeFlagged } = require('./ai');

const ROOT = path.join(__dirname, '..');
const CACHE_FILE = path.join(ROOT, 'flagged-cache.json');
// 공유 사서함(Strata1 등)은 캐시/리포트를 따로 둔다 — 서로의 캐시를 지우지 않도록
function slug(email) { return String(email || '').split('@')[0].replace(/[^A-Za-z0-9_-]/g, '') || 'shared'; }
function cacheFileFor(user) { return user ? path.join(ROOT, `flagged-cache-${slug(user)}.json`) : CACHE_FILE; }
const REPORT_DIR = path.join(ROOT, 'reports');
const LOG_FILE = path.join(ROOT, 'log.txt');

function log(msg) {
  const line = `[${new Date().toLocaleString('sv-SE')}] ${msg}`;
  console.log(line);
  try { fs.appendFileSync(LOG_FILE, line + '\n'); } catch (e) { }
}
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

async function run(cfg, mailbox, folders) {
  // 메일은 이 PC의 Outlook에서 직접 읽는다 (서버 로그인·권한·회사 계정 구분과 무관).
  //   ① --mailbox 로 지정한 사서함(주소 또는 스토어 이름) → ② Outlook 화면에서 보고 있는 사서함 → ③ Outlook 기본 사서함
  // folders: 폴더 spec 배열 ('.'=Inbox, 그 외 = Inbox 하위 경로) — 비어 있으면 Inbox 만
  let target = String(mailbox || '').trim().toLowerCase();
  if (!target) {
    try { target = (await getActiveOutlookAccount()) || ''; } catch (e) { target = ''; }
  }
  const folderSpecs = Array.isArray(folders) ? folders.filter(Boolean) : [];
  log(`🚩 Flagged Summary 생성 시작${target ? ` — 사서함: ${target}` : ' — Outlook 기본 사서함'}`
    + (folderSpecs.length ? ` — 폴더 ${folderSpecs.length}개` : ''));

  let read;
  try {
    read = await readFlagged(target, { max: 200, bodyChars: 6000, folders: folderSpecs });
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
    // classic Outlook 인데도 실패 — 실제 오류를 그대로 보여준다 (권한 수준 차이 등)
    throw new Error('Outlook 은 켜져 있는데 연결하지 못했습니다.\n'
      + '바(또는 Outlook)를 관리자 권한으로 실행했다면 둘 다 일반 권한으로 다시 실행해 주세요.\n'
      + '원인: ' + e.message);
  }
  const boxArg = read.mailbox || target || '';
  const msgs = read.msgs;
  const runFolders = new Set(msgs.map(m => m._folder || 'Inbox'));
  const folderNote = folderSpecs.length > 1 ? ` · folders: ${folderSpecs.length}` : '';
  const accountNote = `${boxArg || '(기본 사서함)'}${folderNote} · Outlook에서 직접 읽음`;
  log(`👤 대상 사서함: ${boxArg || '(기본)'} — flag ${msgs.length}건 `
    + `(방식 ${read.how || '?'}, 검사한 메일 ${read.inboxCount}통, 걸린 ${read.flagCount}건`
    + (folderSpecs.length ? `, 폴더 ${folderSpecs.length}개` : '') + `)`);

  // 오래 방치된 것이 맨 위
  msgs.sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime));

  // 요약 캐시: 받은 메일 본문은 사실상 불변이므로 id 기준으로 한 번만 요약한다
  const cacheFile = cacheFileFor(boxArg);
  let cache = {};
  try { cache = JSON.parse(fs.readFileSync(cacheFile, 'utf8')); } catch (e) { }

  let newCount = 0, failCount = 0;
  for (const m of msgs) {
    // 새 캐시 형식 {en, ko}. 구버전({summary})은 무시하고 다시 요약(EN/KO 둘 다 필요).
    if (cache[m.id]?.en && cache[m.id]?.ko) { m._sum = { en: cache[m.id].en, ko: cache[m.id].ko }; continue; }
    try {
      const fromName = m.from?.emailAddress?.name || m.from?.emailAddress?.address || '(알 수 없음)';
      m._sum = await summarizeFlagged(cfg, {
        fromName, subject: m.subject || '(제목 없음)',
        received: m.receivedDateTime, text: m._text || '',
      });
      cache[m.id] = { en: m._sum.en, ko: m._sum.ko, folder: m._folder || 'Inbox', cachedAt: new Date().toISOString() };
      newCount++;
      log(`🚩 요약 생성(EN+KO): "${m.subject}"`);
    } catch (e) {
      // 실패는 캐시하지 않음 — 다음 실행 때 재시도
      m._sum = null;
      failCount++;
      log(`⚠️  요약 실패 "${m.subject}": ${e.message}`);
    }
  }

  // flag가 해제된 메일의 캐시는 정리 — 단, 이번에 안 읽은 폴더의 캐시는 남겨둔다
  // (일부 폴더만 골라 돌려도 다른 폴더의 요약 캐시가 지워지지 않도록)
  const live = new Set(msgs.map(m => m.id));
  const requested = new Set(folderSpecs.length
    ? folderSpecs.map(s => s === '.' ? 'Inbox' : s.split('\u001f').join('/'))
    : ['Inbox']);
  for (const id of Object.keys(cache)) {
    if (live.has(id)) continue;
    if (requested.has(cache[id]?.folder || 'Inbox')) delete cache[id];
  }
  try { fs.writeFileSync(cacheFile, JSON.stringify(cache, null, 2)); } catch (e) { }

  fs.mkdirSync(REPORT_DIR, { recursive: true });
  const d = new Date(), z = n => String(n).padStart(2, '0');
  const stamp = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}_${z(d.getHours())}${z(d.getMinutes())}`;
  // 하나의 리포트에 영어(왼쪽)+한글(오른쪽) 나란히
  const file = path.join(REPORT_DIR, `Flagged_Summary_${boxArg ? slug(boxArg) + '_' : ''}${stamp}.html`);
  // Unflag 버튼용 독립 도우미 서버를 detached로 띄운다 (감시자와 무관하게 동작, 20분 뒤 자동 종료).
  // EntryID는 사서함과 무관하게 고유하므로 도우미는 하나면 된다
  const _uport = cfg.unflagPort || 3940;
  const _usecret = cfg.unflagSecret || '';
  const _ubase = `http://127.0.0.1:${_uport}`;
  try {
    const child = spawn(process.execPath,
      [path.join(__dirname, 'unflag-server.js'), String(_uport), cfg.clientId || '', cfg.tenant || '', _usecret],
      { detached: true, stdio: 'ignore', windowsHide: true });
    child.unref();
    log(`🚩 Unflag 도우미 시작 (127.0.0.1:${_uport}, 20분 후 자동 종료)`);
  } catch (e) { log(`⚠️  Unflag 도우미 시작 실패: ${e.message}`); }
  // 제목을 누르면 Outlook에서 그 메일이 열리도록 로컬 도우미 링크를 건다
  for (const m of msgs) {
    m.webLink = `${_ubase}/open?id=${encodeURIComponent(m.id)}` + (_usecret ? `&k=${encodeURIComponent(_usecret)}` : '');
  }
  fs.writeFileSync(file, render(msgs, accountNote, _ubase, _usecret));
  log(`🚩 Flagged Summary 완료 (${boxArg} Inbox) — flag ${msgs.length}건, 새 요약 ${newCount}건${failCount ? `, 실패 ${failCount}건` : ''}`);

  // 기본 브라우저로 열기 (영어+한글 한 페이지)
  execFile('cmd.exe', ['/c', 'start', '', file], { windowsHide: true });
}

function ageDays(iso) { return Math.floor((Date.now() - new Date(iso).getTime()) / 86400000); }

const T = {
  title: '🚩 Flagged Email Summary (Inbox)', acct: 'account',
  sortLabel: 'Sort:', sortOld: 'Oldest first', sortNew: 'Newest first',
  colLabel: 'Columns:', colEn: 'English', colKo: 'Korean',
  cTotal: 'Flagged (open)', cOver7: 'over 7 days', cOldest: 'oldest flag',
  none: '🎉 All clear — no flagged emails in the Inbox.',
  fail: 'Summary failed — will retry next run. Preview: ',
  foot: 'Click a subject to open that email in Outlook. Click Unflag to clear the flag. Use the English / Korean buttons to show or hide a language column.',
  noSubj: '(no subject)', unflag: 'Unflag',
  confirmUnflag: 'Unflag this email? It will be removed from the list.',
  unflagOff: 'Could not reach the unflag helper. Re-run Flagged Summary from the bar (it starts the helper), then try again.',
};

function bulletHtml(txt) {
  let o = '';
  for (const raw of String(txt || '').split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    if (/^(ACTION|조치)\s*:/i.test(line)) o += `<div class="act">${esc(line)}</div>`;
    else o += `<div>${esc(line)}</div>`;
  }
  return o;
}

function render(msgs, accountNote, ubase, secret) {
  const t = T;
  const total = msgs.length;
  const oldest = total ? ageDays(msgs[0].receivedDateTime) : 0;
  const over7 = msgs.filter(m => ageDays(m.receivedDateTime) >= 7).length;

  let rows = '';
  for (const m of msgs) {
    const days = ageDays(m.receivedDateTime);
    const badgeCls = days >= 7 ? 'red' : days >= 3 ? 'or' : 'gray';
    const fromName = esc(m.from?.emailAddress?.name || m.from?.emailAddress?.address || '');
    const dt = new Date(m.receivedDateTime).toLocaleString('en-AU', { day: '2-digit', month: 'short', year: 'numeric' });
    const subject = esc(m.subject || t.noSubj);
    const title = m.webLink ? `<a href="${esc(m.webLink)}" target="_blank">${subject}</a>` : subject;
    let body;
    if (m._sum) {
      // 영어(왼쪽) + 한글(오른쪽) 두 칼럼
      body = `<div class="sums">`
        + `<div class="col colen"><div class="clab">${t.colEn}</div><div class="sum">${bulletHtml(m._sum.en)}</div></div>`
        + `<div class="col colko"><div class="clab">한국어</div><div class="sum">${bulletHtml(m._sum.ko)}</div></div>`
        + `</div>`;
    } else {
      body = `<div class="sum fail">${t.fail}${esc((m.bodyPreview || '').slice(0, 150))}</div>`;
    }
    const folderChip = (m._folder && m._folder !== 'Inbox') ? ` · <span class="fold">📁 ${esc(m._folder)}</span>` : '';
    rows += `<div class="mail" data-id="${esc(m.id)}" data-ts="${new Date(m.receivedDateTime).getTime()}">`
      + `<div class="top"><span class="badge ${badgeCls}">${days}d</span><span class="subj">${title}</span>`
      + `<button class="unf" onclick="xUnflag(this)">${t.unflag}</button></div>`
      + `<div class="meta">${fromName} · ${dt}${folderChip}</div>`
      + body
      + `</div>`;
  }
  if (!total) rows = `<div class="none">${t.none}</div>`;

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Flagged Email Summary</title><style>
body{font:14px/1.5 "Segoe UI","Malgun Gothic",Arial,sans-serif;color:#222;margin:0;background:#f2f4f5}
.hd{background:#0F6CBD;color:#fff;padding:16px 24px}
.hd .t{font-size:20px;font-weight:bold}
.hd small{display:block;font-size:12px;opacity:.95;margin-top:3px}
.wrap{max-width:1080px;margin:0 auto;padding:18px 24px}
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
.unf{margin-left:auto;background:#eef2f6;border:1px solid #cfd8e0;color:#33475b;border-radius:6px;padding:3px 12px;font-size:12px;cursor:pointer;white-space:nowrap}
.unf:hover{background:#c1272d;color:#fff;border-color:#c1272d}
.unf:disabled{opacity:.5;cursor:default}
.meta{color:#666;font-size:12px;margin:2px 0 8px}
.fold{color:#0F6CBD;font-weight:bold}
.sums{display:flex;gap:16px}
.col{flex:1;min-width:0}
.colko{border-left:1px solid #e6ebef;padding-left:16px}
.clab{font-size:10px;color:#9aa5b1;font-weight:bold;letter-spacing:.5px;text-transform:uppercase;margin-bottom:3px}
.sum{border-left:3px solid #dbe6ef;padding-left:10px;color:#333}
.sum .act{font-weight:bold;color:#8A4A00;margin-top:4px}
.fail{color:#999;font-style:italic}
body.hide-en .colen{display:none}
body.hide-ko .colko{display:none}
body.hide-en .colko,body.hide-ko .colen{border-left:none;padding-left:0}
body.hide-en .clab,body.hide-ko .clab{display:none}
.none{background:#fff;border:1px solid #d9dee3;border-radius:6px;padding:24px;text-align:center;font-size:16px}
.bar{margin:0 0 12px;color:#555;font-size:13px}
.sortb{background:#fff;border:1px solid #cfd8e0;color:#33475b;border-radius:6px;padding:4px 12px;font-size:13px;cursor:pointer;margin-left:6px}
.sortb.on{background:#0F6CBD;color:#fff;border-color:#0F6CBD}
.sep{display:inline-block;width:22px}
.ft{color:#888;font-size:11px;margin:22px 0}
</style></head><body>
<div class="hd"><div class="t">${t.title}</div><small>${new Date().toLocaleString('en-AU')} · ${t.acct}: ${esc(accountNote)}</small></div>
<div class="wrap">
<div class="sum-cards">
<div class="card"><b id="cnt-total">${total}</b><span>${t.cTotal}</span></div>
<div class="card${over7 ? ' red' : ''}"><b>${over7}</b><span>${t.cOver7}</span></div>
<div class="card"><b>${oldest}d</b><span>${t.cOldest}</span></div>
</div>
${total ? `<div class="bar">${t.sortLabel}<button id="sort-old" class="sortb on" onclick="xSort('old')">${t.sortOld}</button><button id="sort-new" class="sortb" onclick="xSort('new')">${t.sortNew}</button><span class="sep"></span>${t.colLabel}<button id="col-en" class="sortb on" onclick="xCol('en')">${t.colEn}</button><button id="col-ko" class="sortb on" onclick="xCol('ko')">${t.colKo}</button></div>` : ''}
<div id="mails">${rows}</div>
<p class="ft">${t.foot}</p>
</div>
<script>
var UB=${JSON.stringify(ubase || '')},UK=${JSON.stringify(secret || '')};
var CFM=${JSON.stringify(t.confirmUnflag)},OFFM=${JSON.stringify(t.unflagOff)},ULBL=${JSON.stringify(t.unflag)};
function xDec(){var el=document.getElementById('cnt-total');if(el){var n=(parseInt(el.textContent,10)||1)-1;if(n<0)n=0;el.textContent=n;}}
function xSort(dir){var w=document.getElementById('mails');if(!w)return;var cs=[].slice.call(w.querySelectorAll('.mail'));cs.sort(function(a,b){var ta=+a.getAttribute('data-ts'),tb=+b.getAttribute('data-ts');return dir==='new'?tb-ta:ta-tb;});cs.forEach(function(c){w.appendChild(c);});var o=document.getElementById('sort-old'),n=document.getElementById('sort-new');if(o)o.className='sortb'+(dir==='old'?' on':'');if(n)n.className='sortb'+(dir==='new'?' on':'');}
function xCol(which){var b=document.body;b.classList.toggle('hide-'+which);if(b.classList.contains('hide-en')&&b.classList.contains('hide-ko')){b.classList.remove('hide-'+which);return;}['en','ko'].forEach(function(w){var btn=document.getElementById('col-'+w);if(btn)btn.className='sortb'+(b.classList.contains('hide-'+w)?'':' on');});}
function xUnflag(btn){
  var card=btn.closest('.mail');if(!card)return;
  var id=card.getAttribute('data-id');if(!id)return;
  if(!confirm(CFM))return;
  btn.disabled=true;btn.textContent='...';
  var url=UB+'/unflag?id='+encodeURIComponent(id)+(UK?'&k='+encodeURIComponent(UK):'');
  fetch(url).then(function(r){
    if(!r.ok)throw new Error('s'+r.status);
    card.style.transition='opacity .3s';card.style.opacity='0';
    setTimeout(function(){card.remove();xDec();},300);
  }).catch(function(){btn.disabled=false;btn.textContent=ULBL;alert(OFFM);});
}
</script>
</body></html>`;
}

module.exports = { run };
