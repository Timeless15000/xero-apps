// OUTLOOK Bar 용 메일 읽기 — Microsoft Graph 로 직접 읽는다 (COM 자동화 없음)
// why: 예전 outlook-read.js 는 실행 중인 클래식 Outlook(COM)에서만 읽을 수 있어
//      새 Outlook(olk.exe)·Outlook 꺼진 PC 에서는 바의 버튼이 아무것도 못 했다.
//      Graph 는 Outlook 종류·실행 여부와 무관하고, 감시자와 같은 로그인(tokens.json)을 그대로 쓴다.
// 반환 모양은 outlook-read.js 와 같게 맞춰 flagged.js / review-daily.js 가 그대로 쓴다.
const { getAccessToken } = require('./auth');

const BASE = 'https://graph.microsoft.com/v1.0';
const SEP = '\u001f';                       // 폴더 spec 구분자 (바 ↔ 프로그램, outlook-read.js 와 동일)
const VERB_PROP = 'Integer 0x1081';          // PR_LAST_VERB_EXECUTED: 102=Reply 103=ReplyAll 104=Forward

function mbox(user) { return user ? `/users/${encodeURIComponent(user)}` : '/me'; }
function odataStr(s) { return String(s).replace(/'/g, "''"); }
function lower(s) { return String(s || '').trim().toLowerCase(); }
function isSmtp(s) { return /^[^\s@]+@[^\s@]+$/.test(String(s || '')); }

// 테스트에서 fetch/토큰을 바꿔 끼울 수 있게 팩토리로 만든다
function makeGraphReader({ getToken, fetchImpl = fetch } = {}) {
  const token = getToken || (clientId => getAccessToken(clientId));

  async function gf(clientId, url, { headers = {}, method = 'GET', body } = {}, retry = 2) {
    const tk = await token(clientId);
    let res;
    try {
      res = await fetchImpl(url.startsWith('http') ? url : BASE + url, {
        method, body,
        signal: AbortSignal.timeout(40000),
        headers: { authorization: 'Bearer ' + tk, 'content-type': 'application/json', ...headers },
      });
    } catch (e) {
      if (retry > 0) { await new Promise(r => setTimeout(r, 3000)); return gf(clientId, url, { headers, method, body }, retry - 1); }
      throw e;
    }
    if (res.status === 429 && retry > 0) {
      const wait = (parseInt(res.headers.get('retry-after')) || 10) * 1000;
      await new Promise(r => setTimeout(r, wait));
      return gf(clientId, url, { headers, method, body }, retry - 1);
    }
    if (res.status === 204) return null;
    const data = await res.json().catch(() => null);
    if (!res.ok) {
      const msg = data?.error?.message || res.statusText || '';
      if (res.status === 404 && /mailbox|user/i.test(msg)) throw new Error(`MAILBOX_NOT_FOUND: ${msg}`);
      if (res.status === 403) throw new Error(`NO_ACCESS: ${msg}`);
      throw new Error(`Graph ${res.status}: ${msg}`);
    }
    return data;
  }

  // 페이지를 따라가며 최대 max 건 수집
  async function collect(clientId, url, max, headers) {
    const all = [];
    let next = url;
    while (next && all.length < max) {
      const d = await gf(clientId, next, { headers });
      all.push(...(d?.value || []));
      next = d?.['@odata.nextLink'] || null;
    }
    return all.slice(0, max);
  }

  // 로그인한 본인 — { smtp, name, aliases[] }
  async function me(clientId) {
    const m = await gf(clientId, '/me?$select=displayName,mail,userPrincipalName,proxyAddresses');
    const smtp = lower(m.mail || m.userPrincipalName);
    const aliases = new Set([smtp]);
    for (const p of m.proxyAddresses || []) {
      const s = lower(String(p).replace(/^smtp:/i, ''));
      if (isSmtp(s)) aliases.add(s);
    }
    return { smtp, name: m.displayName || smtp, aliases: [...aliases].filter(Boolean) };
  }

  // 바의 사서함 메뉴용 목록 — 본인 + config.json 에 적힌 공유 사서함
  //   cfg.barMailboxes: ["strata@...", ...]  (있으면 이것만)
  //   없으면 cfg.pinSync.accounts[*].mailboxes 중 같은 테넌트(cfg.tenant) 것
  // 공유 사서함은 Graph 가 "내가 접근 가능한 사서함 목록"을 주지 않아서 설정에서 읽는다.
  function configuredMailboxes(cfg) {
    const out = [];
    if (Array.isArray(cfg?.barMailboxes) && cfg.barMailboxes.length) {
      for (const x of cfg.barMailboxes) {
        const smtp = lower(typeof x === 'string' ? x : x?.smtp);
        if (isSmtp(smtp)) out.push({ smtp, name: (typeof x === 'object' && x?.name) || smtp });
      }
      return out;
    }
    const tenant = lower(cfg?.tenant);
    for (const a of cfg?.pinSync?.accounts || []) {
      if (tenant && a?.tenant && lower(a.tenant) !== tenant) continue;
      for (const s of a?.mailboxes || []) if (isSmtp(lower(s))) out.push({ smtp: lower(s), name: lower(s) });
    }
    return out;
  }

  async function listMailboxes(clientId, cfg, extra = []) {
    const my = await me(clientId);
    const seen = new Set();
    const list = [];
    const add = b => {
      const smtp = lower(b?.smtp);
      if (!isSmtp(smtp) || seen.has(smtp)) return;
      seen.add(smtp);
      list.push({ smtp, name: b.name || smtp });
    };
    add(my);
    for (const b of configuredMailboxes(cfg)) add(b);
    for (const b of extra || []) add(b);          // 클래식 Outlook 이 켜져 있으면 거기 붙은 사서함도 합친다
    return { me: my, list };
  }

  // Inbox 하위 폴더 이름으로 폴더 id 를 찾는다 (spec: '.' 또는 'A\u001fB')
  async function resolveFolder(clientId, user, spec) {
    let id = 'inbox';
    if (!spec || spec === '.') return id;
    for (const name of spec.split(SEP)) {
      const d = await gf(clientId, `${mbox(user)}/mailFolders/${encodeURIComponent(id)}/childFolders?$filter=displayName eq '${odataStr(name)}'&$select=id,displayName&$top=5`);
      const hit = (d?.value || []).find(f => lower(f.displayName) === lower(name)) || (d?.value || [])[0];
      if (!hit) throw new Error(`FOLDER_NOT_FOUND: ${spec.split(SEP).join(' / ')}`);
      id = hit.id;
    }
    return id;
  }

  // 바의 폴더 선택 창용 — Inbox + 하위 폴더 (3단계까지) [{ spec, display }]
  async function listFolders(clientId, user) {
    const out = [{ spec: '.', display: 'Inbox' }];
    async function walk(id, spec, disp, depth) {
      let kids = [];
      try { kids = await collect(clientId, `${mbox(user)}/mailFolders/${encodeURIComponent(id)}/childFolders?$select=id,displayName,childFolderCount&$top=100`, 300); }
      catch (e) { return; }
      kids.sort((a, b) => String(a.displayName || '').localeCompare(String(b.displayName || ''), 'en'));
      for (const c of kids) {
        const nm = c.displayName || '';
        if (!nm) continue;
        const ns = spec === '.' ? nm : spec + SEP + nm;
        const nd = disp + ' / ' + nm;
        out.push({ spec: ns, display: nd });
        if (depth < 3 && (c.childFolderCount || 0) > 0) await walk(c.id, ns, nd, depth + 1);
      }
    }
    await walk('inbox', '.', 'Inbox', 1);
    return out;
  }

  function fromOf(m) {
    const ea = m.from?.emailAddress || m.sender?.emailAddress || {};
    return { name: ea.name || '', addr: lower(ea.address) };
  }

  // flag 된 메일 — outlook-read.readFlagged 와 같은 모양
  async function readFlagged(clientId, mailbox, { max = 200, bodyChars = 6000, folders = [] } = {}) {
    const user = isSmtp(mailbox) ? lower(mailbox) : '';
    const specs = (folders || []).filter(Boolean);
    if (!specs.length) specs.push('.');
    const msgs = [];
    const prefer = { prefer: 'outlook.body-content-type="text"' };
    for (const spec of specs) {
      const fid = await resolveFolder(clientId, user, spec);
      const display = spec === '.' ? 'Inbox' : spec.split(SEP).join('/');
      const url = `${mbox(user)}/mailFolders/${encodeURIComponent(fid)}/messages?$filter=flag/flagStatus eq 'flagged'&$top=50`
        + `&$select=id,subject,from,sender,receivedDateTime,webLink,bodyPreview,body`;
      const items = await collect(clientId, url, Math.max(1, max - msgs.length), prefer);
      for (const it of items) {
        const f = fromOf(it);
        const text = String(it.body?.content || it.bodyPreview || '').replace(/\r/g, '').trim().slice(0, bodyChars);
        msgs.push({
          id: it.id,
          subject: it.subject || '',
          from: { emailAddress: { name: f.name, address: f.addr } },
          receivedDateTime: it.receivedDateTime || new Date(0).toISOString(),
          bodyPreview: String(it.bodyPreview || text).replace(/\s+/g, ' ').slice(0, 200),
          webLink: it.webLink || '',
          _text: text,
          _folder: display,
        });
      }
      if (msgs.length >= max) break;
    }
    let box = user;
    if (!box) { try { box = (await me(clientId)).smtp; } catch (e) { box = ''; } }
    return { mailbox: box, msgs, how: 'graph', inboxCount: msgs.length, flagCount: msgs.length };
  }

  // 최근 hours 시간 Inbox 메일 + 그 기간 보낸 메일의 대화 id — outlook-read.readRecent 와 같은 모양
  async function readRecent(clientId, mailbox, { hours = 24, bodyChars = 700, max = 400 } = {}) {
    const user = isSmtp(mailbox) ? lower(mailbox) : '';
    const since = new Date(Date.now() - hours * 3600000).toISOString();
    const inboxUrl = `${mbox(user)}/mailFolders/inbox/messages?$filter=receivedDateTime ge ${since}`
      + `&$orderby=receivedDateTime desc&$top=50`
      + `&$select=id,subject,from,sender,receivedDateTime,bodyPreview,isRead,conversationId,webLink`
      + `&$expand=singleValueExtendedProperties($filter=id eq '${VERB_PROP}')`;
    const sentUrl = `${mbox(user)}/mailFolders/sentitems/messages?$filter=sentDateTime ge ${since}&$top=50&$select=id,conversationId`;
    const [inbox, sent, my] = await Promise.all([
      collect(clientId, inboxUrl, max),
      collect(clientId, sentUrl, 500).catch(() => []),
      me(clientId).catch(() => ({ smtp: '', aliases: [] })),
    ]);
    const myAddrs = new Set([...(my.aliases || []), user].filter(Boolean));
    const items = inbox.map(m => {
      const f = fromOf(m);
      let verb = 0;
      for (const p of m.singleValueExtendedProperties || []) if (String(p.id).toLowerCase() === VERB_PROP.toLowerCase()) verb = parseInt(p.value, 10) || 0;
      return {
        id: m.id,
        subject: m.subject || '',
        name: f.name,
        addr: f.addr,
        received: m.receivedDateTime || new Date(0).toISOString(),
        preview: String(m.bodyPreview || '').slice(0, bodyChars),
        unread: m.isRead === false,
        verb,
        conv: m.conversationId || '',
        mine: !!f.addr && myAddrs.has(f.addr),
        webLink: m.webLink || '',
      };
    });
    return {
      mailbox: user || my.smtp || '',
      items,
      sentSet: new Set(sent.map(s => s.conversationId).filter(Boolean)),
      myAddrs: [...myAddrs],
      scanned: inbox.length,
    };
  }

  async function unflag(clientId, id, mailbox) {
    const user = isSmtp(mailbox) ? lower(mailbox) : '';
    await gf(clientId, `${mbox(user)}/messages/${encodeURIComponent(id)}`, {
      method: 'PATCH', body: JSON.stringify({ flag: { flagStatus: 'notFlagged' } }),
    });
    return 'unflagged';
  }

  // Graph 메일 id → 클래식 Outlook EntryID (클래식이 켜져 있을 때 그 창에서 열어주려고)
  async function toEntryId(clientId, id, mailbox) {
    const user = isSmtp(mailbox) ? lower(mailbox) : '';
    const d = await gf(clientId, `${mbox(user)}/translateExchangeIds`, {
      method: 'POST',
      body: JSON.stringify({ inputIds: [id], sourceIdType: 'restId', targetIdType: 'entryId' }),
    });
    return d?.value?.[0]?.targetId || '';
  }

  return { me, listMailboxes, configuredMailboxes, listFolders, resolveFolder, readFlagged, readRecent, unflag, toEntryId };
}

const def = makeGraphReader();

// 로그인이 되어 있는지 (tokens.json + 갱신 가능) — 바가 버튼을 누르기 전에 확인한다
async function checkLogin(clientId) {
  try { await getAccessToken(clientId); return true; } catch (e) { return false; }
}

// 오류를 직원이 알아볼 문장으로
function explain(e) {
  const m = String(e?.message || e || '');
  if (/로그인이 필요|토큰 갱신 실패|npm run login/.test(m)) return 'LOGIN_REQUIRED';
  if (/^MAILBOX_NOT_FOUND/.test(m)) return 'MAILBOX_NOT_FOUND';
  if (/^NO_ACCESS/.test(m)) return 'NO_ACCESS';
  if (/^FOLDER_NOT_FOUND/.test(m)) return 'FOLDER_NOT_FOUND';
  if (/fetch failed|ENOTFOUND|ECONN|ETIMEDOUT|timeout|abort/i.test(m)) return 'NETWORK';
  return '';
}

module.exports = { makeGraphReader, checkLogin, explain, SEP, ...def };
