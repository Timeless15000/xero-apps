// Microsoft ID 플랫폼 — 디바이스 코드 로그인 + 토큰 자동 갱신 (외부 라이브러리 없음)
// config.json의 tenant: 개인 계정은 "consumers", 회사 계정은 테넌트 도메인(예: "timelesscommercial.com.au")
const fs = require('fs');
const path = require('path');

const TOKEN_FILE = path.join(__dirname, '..', 'tokens.json');
const SCOPE = 'offline_access User.Read Mail.ReadWrite Mail.ReadWrite.Shared';

let AUTH_BASE = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
function setTenant(tenant) {
  if (tenant) AUTH_BASE = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0`;
}

function loadTokens() {
  try { return JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8')); } catch (e) { return null; }
}
function saveTokens(t) {
  fs.writeFileSync(TOKEN_FILE, JSON.stringify(t, null, 2));
}

async function tokenRequest(params) {
  const res = await fetch(AUTH_BASE + '/token', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(params).toString(),
  });
  return { ok: res.ok, data: await res.json() };
}

// 최초 1회: 디바이스 코드 로그인 (화면의 코드를 microsoft.com/devicelogin 에 입력)
async function login(clientId) {
  const res = await fetch(AUTH_BASE + '/devicecode', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: clientId, scope: SCOPE }).toString(),
  });
  const dc = await res.json();
  if (!dc.device_code) throw new Error('디바이스 코드 발급 실패: ' + JSON.stringify(dc));
  console.log('\n════════════════════════════════════════════════');
  console.log(dc.message); // "To sign in, use a web browser to open ... and enter the code ..."
  console.log('════════════════════════════════════════════════\n');
  const interval = (dc.interval || 5) * 1000;
  while (true) {
    await new Promise(r => setTimeout(r, interval));
    const { data } = await tokenRequest({
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: clientId,
      device_code: dc.device_code,
    });
    if (data.access_token) {
      saveTokens({
        access_token: data.access_token,
        refresh_token: data.refresh_token,
        expires_at: Date.now() + (data.expires_in - 120) * 1000,
      });
      console.log('✅ 로그인 성공 — 토큰이 tokens.json에 저장되었습니다.');
      return;
    }
    if (data.error === 'authorization_pending') continue;
    if (data.error === 'slow_down') { await new Promise(r => setTimeout(r, 5000)); continue; }
    throw new Error('로그인 실패: ' + (data.error_description || data.error));
  }
}

// 유효한 액세스 토큰 반환 (만료되면 refresh_token으로 자동 갱신)
async function getAccessToken(clientId) {
  let t = loadTokens();
  if (!t) throw new Error('로그인이 필요합니다. 먼저 실행하세요:  npm run login');
  if (Date.now() < t.expires_at) return t.access_token;
  const { ok, data } = await tokenRequest({
    grant_type: 'refresh_token',
    client_id: clientId,
    refresh_token: t.refresh_token,
    scope: SCOPE,
  });
  if (!ok || !data.access_token) {
    throw new Error('토큰 갱신 실패 — 다시 로그인하세요 (npm run login): ' + (data.error_description || JSON.stringify(data)));
  }
  saveTokens({
    access_token: data.access_token,
    refresh_token: data.refresh_token || t.refresh_token,
    expires_at: Date.now() + (data.expires_in - 120) * 1000,
  });
  return data.access_token;
}

module.exports = { login, getAccessToken, setTenant };
