// Claude API — 답장 필요 여부 분류 + 답장 초안 생성
async function callClaude(apiKey, model, system, user, maxTokens) {
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model, max_tokens: maxTokens, system,
      messages: [{ role: 'user', content: user }],
    }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(`Claude ${res.status}: ${data?.error?.message || 'API 오류'}`);
  return (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
}

// 1단계: 답장이 필요한 메일인지 분류
async function classify(cfg, msg) {
  const system = `You classify incoming emails for a small service business owner.
Decide if the email needs a personal reply from the owner.

needsReply = true: customer inquiries, booking/quote requests, schedule changes, complaints, questions from clients or business partners, personal messages that expect an answer.
needsReply = false: newsletters, ads, automated notifications, receipts, shipping updates, social media alerts, spam, no-reply senders, FYI-only messages.

IMPORTANT: If the email is sent by an automated system (utility company, bank, government, platform notification), needsReply = false even when it asks the recipient to take action (e.g. "call us to reschedule", "log in to confirm"). An email reply would go to an unmonitored mailbox — the action happens by phone or website, not by replying.

Categories: "inquiry" (question/booking/quote), "complaint" (dissatisfaction/problem), "business" (partner/staff/admin), "thanks" (positive feedback), "other".

Respond with ONLY a JSON object, no other text:
{"needsReply": true/false, "category": "...", "language": "en"/"ko", "reason": "one short sentence"}`;
  const user = `From: ${msg.fromName} <${msg.fromEmail}>`
    + (msg.replyToEmail && msg.replyToEmail !== msg.fromEmail ? `\nReply-To: ${msg.replyToEmail}` : '')
    + `\nSubject: ${msg.subject}
Preview: ${msg.preview}`;
  const out = await callClaude(cfg.anthropicApiKey, cfg.classifyModel, system, user, 300);
  const m = out.match(/\{[\s\S]*\}/);
  if (!m) throw new Error('분류 결과 파싱 실패: ' + out.slice(0, 200));
  return JSON.parse(m[0]);
}

// 2단계: 톤 가이드에 맞춘 답장 초안 생성 (본문 텍스트만 반환)
async function draftReply(cfg, toneGuide, msg, category, opts = {}) {
  const internalNote = opts.isInternal
    ? `- This email is from INTERNAL STAFF (same company domain). If it is written in Korean, you MUST reply in Korean following the "내부 직원 메일" section of the tone guide: 존댓말, no greetings/thanks, answer first (결론부터), short and directive.`
    : '';
  const system = `You write reply drafts for the owner of a small service business.
Follow this tone guide EXACTLY — it is the single source of truth for style, structure, and signature:

${toneGuide}

Rules:
- Write ONLY the reply body text (no subject line, no explanations, no markdown).
- Reply in the same language as the incoming email (default: English). Detected language: ${opts.language || 'unknown'}.
${internalNote}
- Category of this email: ${category}. Apply the matching tone from the guide.
- If any fact is uncertain (price, date, availability), write [CHECK: ...] instead of inventing it.
- Do NOT end with a sign-off or name (no "Kind regards," no signature block) — the owner's real signature is appended automatically after your text.
- Keep it short and human. This draft will be reviewed by the owner before sending.`;
  const user = `Incoming email:
From: ${msg.fromName} <${msg.fromEmail}>
Subject: ${msg.subject}

${msg.text}`;
  return (await callClaude(cfg.anthropicApiKey, cfg.draftModel, system, user, 1000)).trim();
}

// 긴 메일(스레드)을 짧은 영어 불릿으로 요약 — 답장 쓰기 전 참고용
async function summarize(cfg, msg) {
  const system = `You are the assistant to Brian, who runs a cleaning and gardening business (Timeless) in Australia.
Summarise a long email thread so Brian can grasp it in about ten seconds before replying.

Rules:
- Output 3-6 bullets, each on its own line starting with "- ".
- Maximum 20 words per bullet. Use fragments, not full sentences.
- Reconstruct thread events in chronological order.
- Keep names, unit/room numbers, SP numbers, order and quote numbers, amounts and dates exactly as written.
- The final line must start with "ACTION: " and state precisely what Brian has to do or decide next.
- No headings, no preamble, no closing remarks. Output only the bullets and the ACTION line.`;
  const user = `From: ${msg.fromName}
Subject: ${msg.subject}

${msg.text}`;
  return (await callClaude(cfg.anthropicApiKey, cfg.summaryModel || cfg.draftModel, system, user, 400)).trim();
}

// flag된 메일(마무리 안 된 일) 1건을 아주 짧게 요약 — 영어 + 한국어 두 버전을 한 번의 호출로
// 반환: { en, ko }  (각각 불릿 1~2줄 + ACTION/조치 한 줄, 예전보다 약 절반 길이)
async function summarizeFlagged(cfg, msg) {
  const system = `You are the assistant to Brian, who runs a cleaning and gardening business (Timeless) in Australia.
The email below was FLAGGED by Brian as unfinished business. Summarise it VERY briefly so he instantly recalls it and what is still outstanding.

Rules:
- Be very concise (about half the length of a normal summary).
- 1-2 short bullets only, each on its own line starting with "- ", max 12 words each. Fragments, not sentences.
- Keep names, unit/room numbers, SP numbers, amounts and dates exactly as written.
- Then one action line: the single next step to finish this.
- Give the summary in BOTH English and Korean.
- Output EXACTLY this structure, nothing else (no headings, no preamble):
EN:
- <bullet>
ACTION: <next step>
KO:
- <불릿>
조치: <다음 할 일>`;
  const user = `From: ${msg.fromName}
Subject: ${msg.subject}
Received: ${msg.received}

${msg.text}`;
  const out = (await callClaude(cfg.anthropicApiKey, cfg.classifyModel, system, user, 500)).trim();
  const enM = out.match(/EN\s*:\s*([\s\S]*?)(?:\n\s*KO\s*:|$)/i);
  const koM = out.match(/KO\s*:\s*([\s\S]*)$/i);
  const en = (enM ? enM[1] : out).trim();
  const ko = (koM ? koM[1] : '').trim();
  return { en: en, ko: ko || en };
}

module.exports = { classify, draftReply, summarize, summarizeFlagged };
