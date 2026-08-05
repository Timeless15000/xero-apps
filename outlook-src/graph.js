// Microsoft Graph API 래퍼 — 메일 조회 · 답장 초안 생성
const { getAccessToken } = require('./auth');

const BASE = 'https://graph.microsoft.com/v1.0';

async function graphFetch(clientId, url, options = {}, retry = 2) {
  const token = await getAccessToken(clientId);
  let res;
  try {
    res = await fetch(url.startsWith('http') ? url : BASE + url, {
      ...options,
      signal: AbortSignal.timeout(30000), // 30초 무응답이면 끊고 재시도
      headers: {
        authorization: 'Bearer ' + token,
        'content-type': 'application/json',
        ...(options.headers || {}),
      },
    });
  } catch (e) { // 일시적 네트워크 오류 — 5초 뒤 재시도
    if (retry > 0) {
      await new Promise(r => setTimeout(r, 5000));
      return graphFetch(clientId, url, options, retry - 1);
    }
    throw e;
  }
  if (res.status === 429 && retry > 0) { // 속도 제한 — 잠시 후 재시도
    const wait = (parseInt(res.headers.get('retry-after')) || 10) * 1000;
    await new Promise(r => setTimeout(r, wait));
    return graphFetch(clientId, url, options, retry - 1);
  }
  if (res.status === 204) return null;
  const data = await res.json().catch(() => null);
  if (!res.ok) throw new Error(`Graph ${res.status}: ${data?.error?.message || res.statusText}`);
  return data;
}

// 대상 사서함 경로. user가 없으면 로그인한 본인(/me), 있으면 공유 사서함(/users/<smtp>).
// 공유 사서함을 읽으려면 Mail.ReadWrite.Shared 권한 + Outlook에서 그 사서함에 접근 권한이 있어야 한다.
function mbox(user) { return user ? `/users/${encodeURIComponent(user)}` : '/me'; }

// 내 계정 정보 (이메일 주소 확인용)
async function getMe(clientId) {
  return graphFetch(clientId, '/me?$select=displayName,mail,userPrincipalName,proxyAddresses');
}

// 특정 시각 이후 받은편지함 메일 목록 (최신순)
// PC가 오래 꺼져 있던 뒤에도 놓치지 않도록 다음 페이지를 따라가며 최대 200통까지 수집
async function getNewMessages(clientId, sinceIso) {
  const q = `/me/mailFolders/inbox/messages?$filter=receivedDateTime gt ${sinceIso}` +
    `&$orderby=receivedDateTime desc&$top=50` +
    `&$select=id,subject,from,replyTo,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isDraft,conversationId`;
  const all = [];
  let url = q;
  while (url && all.length < 200) {
    const data = await graphFetch(clientId, url);
    all.push(...(data.value || []));
    url = data['@odata.nextLink'] || null;
  }
  return all;
}

// flag된 메일 조회 — INBOX 폴더에 직접 있는 메일만 (하위폴더·Archive·Sent·Deleted 제외)
// 이전엔 /me/messages(전체 폴더)라 싱크 잔재까지 200통 넘게 잡혔음. 지금 화면 Inbox 것만 나오도록 inbox로 한정.
// $orderby를 flag 필터와 같이 쓰면 Graph가 거부하는 경우가 있어 정렬은 호출한 쪽에서 한다
async function getFlaggedMessages(clientId, user) {
  const q = `${mbox(user)}/mailFolders/inbox/messages?$filter=flag/flagStatus eq 'flagged'&$top=50` +
    `&$select=id,subject,from,receivedDateTime,webLink,bodyPreview`;
  const all = [];
  let url = q;
  while (url && all.length < 200) {
    const data = await graphFetch(clientId, url);
    all.push(...(data.value || []));
    url = data['@odata.nextLink'] || null;
  }
  return all;
}

// 진단용: flag된 메일을 폴더별로 집계 (싱크 잔재가 어느 폴더에 몇 개 있는지 확인)
async function getFlaggedByFolder(clientId) {
  const folders = {};
  const data = await graphFetch(clientId,
    `/me/messages?$filter=flag/flagStatus eq 'flagged'&$top=100&$select=id,parentFolderId,subject,receivedDateTime`);
  let items = data.value || [];
  let url = data['@odata.nextLink'];
  while (url && items.length < 500) {
    const d = await graphFetch(clientId, url);
    items.push(...(d.value || []));
    url = d['@odata.nextLink'] || null;
  }
  const nameCache = {};
  for (const it of items) {
    const fid = it.parentFolderId;
    if (!nameCache[fid]) {
      try { const f = await graphFetch(clientId, `/me/mailFolders/${fid}?$select=displayName`); nameCache[fid] = f.displayName || fid; }
      catch (e) { nameCache[fid] = fid; }
    }
    const nm = nameCache[fid];
    (folders[nm] = folders[nm] || []).push(it);
  }
  return folders;
}

// flag 해제 (Flagged Summary 리포트의 Unflag 버튼용). status 'notFlagged'로 PATCH.
async function unflagMessage(clientId, id, user) {
  await graphFetch(clientId, `${mbox(user)}/messages/${id}`, {
    method: 'PATCH',
    body: JSON.stringify({ flag: { flagStatus: 'notFlagged' } }),
  });
  return 'unflagged';
}

// 메일 본문 전체 (텍스트로 변환)
async function getMessageText(clientId, id, maxChars = 6000, user) {
  const m = await graphFetch(clientId, `${mbox(user)}/messages/${id}?$select=subject,from,body,toRecipients`);
  let text = m.body?.content || '';
  if ((m.body?.contentType || '').toLowerCase() === 'html') {
    text = text
      .replace(/<style[\s\S]*?<\/style>/gi, '').replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<\/div>/gi, '\n')
      .replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#39;/g, "'").replace(/&quot;/g, '"')
      .replace(/\n{3,}/g, '\n\n').trim();
  }
  return { subject: m.subject, from: m.from, text: text.slice(0, maxChars) };
}

// ── 한글 요약 블록: 받은 메일 본문 맨 위에 삽입 / 일괄 제거 ──
// 앵커(<a name>)는 Outlook/Exchange가 본문을 재가공해도 살아남으므로 마커로 쓴다 (서명의 _MailAutoSig와 같은 원리)
const SUM_START = '<a name="aiKoSummary"></a>';
const SUM_END = '<a name="aiKoSummaryEnd"></a>';
const SUM_TEXT_START = '===== AI 요약 =====';
const SUM_TEXT_END = '===== /AI 요약 =====';

function hasSummary(content) {
  return content.includes('aiKoSummary') || content.includes(SUM_TEXT_START);
}

// 본문에서 요약 블록만 제거 (HTML 마커·텍스트 마커 모두)
function stripSummaryBlocks(content) {
  return content
    .replace(/<a[^>]*name="aiKoSummary"[\s\S]*?name="aiKoSummaryEnd"[^>]*>(?:\s*<\/a>)?/g, '')
    // Outlook이 답장 본문을 재가공하며 남기는 고아 마커(짝 잃은 앵커)도 제거 — 안 지우면 hasSummary가 계속 참이 되어 무한 재시도
    .replace(/<a[^>]*name="?aiKoSummary(?:End)?"?[^>]*>(?:\s*<\/a>)?/g, '')
    .replace(/===== AI 요약 =====[\s\S]*?===== \/AI 요약 =====\n*/g, '');
}

// 요약 삽입. 이미 요약이 있으면 false 반환
// 초안에는 기본적으로 삽입하지 않음 — 자동 답장 초안에 의도적으로 넣을 때만 allowDraft: true
// (초안의 요약은 클래식 Outlook 매크로가 보내기 직전 자동 삭제한다)
async function prependSummary(clientId, id, summaryHtml, summaryText, opts = {}) {
  const m = await graphFetch(clientId, `/me/messages/${id}?$select=body,isDraft`);
  if (m.isDraft && !opts.allowDraft) return false;
  const type = (m.body?.contentType || '').toLowerCase();
  let content = m.body?.content || '';
  if (hasSummary(content)) return false;
  if (type === 'html') {
    const block = SUM_START + summaryHtml + SUM_END;
    const bodyTag = content.match(/<body[^>]*>/i);
    content = bodyTag ? content.replace(bodyTag[0], bodyTag[0] + block) : block + content;
  } else {
    content = `${SUM_TEXT_START}\n${summaryText}\n${SUM_TEXT_END}\n\n` + content;
  }
  await graphFetch(clientId, `/me/messages/${id}`, {
    method: 'PATCH',
    body: JSON.stringify({ body: { contentType: type === 'html' ? 'html' : 'text', content } }),
  });
  return true;
}

// 요약 제거. 'removed' | 'none' 반환 (메일이 없으면 graphFetch가 404를 던짐)
async function removeSummary(clientId, id) {
  const m = await graphFetch(clientId, `/me/messages/${id}?$select=body`);
  const type = (m.body?.contentType || '').toLowerCase();
  const content = m.body?.content || '';
  if (!hasSummary(content)) return 'none';
  const stripped = stripSummaryBlocks(content);
  await graphFetch(clientId, `/me/messages/${id}`, {
    method: 'PATCH',
    body: JSON.stringify({ body: { contentType: type === 'html' ? 'html' : 'text', content: stripped } }),
  });
  return 'removed';
}

// 요약이 삽입된 메일 검색 (state 목록에서 빠졌거나 폴더 이동으로 id가 바뀐 것까지 회수)
async function findSummarizedMessages(clientId) {
  // 현재 헤더("AI Thread Summary")와 구버전 헤더("한글 요약")를 모두 찾는다 —
  // 요약 문구를 또 바꾸면 이 검색어도 같이 바꿔야 회수가 된다
  const q = encodeURIComponent('"AI Thread Summary" OR "한글 요약"');
  const data = await graphFetch(clientId, `/me/messages?$search=${q}&$select=id,subject&$top=100`);
  return data.value || [];
}

// 답장 초안 생성: createReplyAll로 원문의 To/CC 전원이 포함된 초안을 만들고, 위에 우리 답장을 얹는다
// inlineImages: [{ name, contentId, contentBytes }] — 서명 로고 등 본문에 삽입되는 이미지(cid: 참조)
async function createReplyDraft(clientId, messageId, replyHtml, inlineImages = []) {
  // Prefer 헤더로 인용 헤더(Sent: ...)의 시간대를 시드니 기준으로 생성
  const draft = await graphFetch(clientId, `/me/messages/${messageId}/createReplyAll`, {
    method: 'POST', body: '{}',
    headers: { prefer: 'outlook.timezone="AUS Eastern Standard Time"' },
  });
  // 원문에 한글 요약이 삽입돼 있어도 답장 인용문에는 절대 포함되지 않게 제거
  const quoted = stripSummaryBlocks(draft.body?.content || '');
  const merged = (draft.body?.contentType || '').toLowerCase() === 'html'
    ? `<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">${replyHtml}</div>` + quoted
    : replyHtml.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '') + '\n\n' + quoted;
  await graphFetch(clientId, `/me/messages/${draft.id}`, {
    method: 'PATCH',
    body: JSON.stringify({ body: { contentType: 'html', content: merged } }),
  });
  for (const img of inlineImages) {
    await graphFetch(clientId, `/me/messages/${draft.id}/attachments`, {
      method: 'POST',
      body: JSON.stringify({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: img.name,
        contentBytes: img.contentBytes,
        contentType: 'image/png',
        isInline: true,
        contentId: img.contentId,
      }),
    });
  }
  return draft.id;
}

module.exports = {
  getMe, getNewMessages, getMessageText, createReplyDraft,
  prependSummary, removeSummary, findSummarizedMessages, getFlaggedMessages, getFlaggedByFolder, unflagMessage,
};
