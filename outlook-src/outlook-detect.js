// 실행 중인 클래식 Outlook에서 "지금 화면에 보이는 폴더"가 속한 계정(SMTP 주소)을 감지
// Outlook 미실행 · COM 실패 · PST처럼 계정 없는 스토어 · 타임아웃 → null 반환 (에러 아님)
const { execFile } = require('child_process');

// 순서: ①계정 기본 사서함 대조 ②스토어 이름이 주소면 그대로
// ③공유·추가 사서함이면 스토어 이름(또는 사서함 소유자 이름)을 주소록에서 해석해 SMTP를 얻는다
const PS_SCRIPT = `
function Get-OutlookApp {
  $app = $null
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $app = $null }
  if ($null -eq $app) {
    # ROT 등록이 안 된 경우(권한 수준 차이 등)에도 붙을 수 있게 COM 개체로 직접 연결
    try { $app = New-Object -ComObject Outlook.Application } catch { $app = $null }
  }
  return $app
}
try {
  $ol = Get-OutlookApp
  if ($null -eq $ol) { throw 'NO_OUTLOOK' }
  $exp = $ol.ActiveExplorer()
  if ($exp -and $exp.CurrentFolder -and $exp.CurrentFolder.Store) {
    $store = $exp.CurrentFolder.Store

    # 1) 계정의 기본 사서함(DeliveryStore)과 StoreID가 같으면 그 계정 주소
    foreach ($a in $ol.Session.Accounts) {
      if ($a.DeliveryStore -and $a.DeliveryStore.StoreID -eq $store.StoreID) { Write-Output $a.SmtpAddress; exit }
    }

    # 2) 스토어 표시 이름 자체가 메일 주소면 그대로 사용
    if ($store.DisplayName -match '@') { Write-Output $store.DisplayName; exit }

    # 3) 공유 사서함 / 추가로 붙인 사서함(예: "Strata1"): 이름을 주소록에서 해석해 SMTP를 얻는다
    $names = @($store.DisplayName)
    try { $names += $store.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x661B001F') } catch { }
    foreach ($n in $names) {
      if (-not $n) { continue }
      try {
        $r = $ol.Session.CreateRecipient($n)
        $null = $r.Resolve()
        if ($r.Resolved) {
          try { $eu = $r.AddressEntry.GetExchangeUser() } catch { $eu = $null }
          if ($eu -and $eu.PrimarySmtpAddress) { Write-Output $eu.PrimarySmtpAddress; exit }
          try { $dl = $r.AddressEntry.GetExchangeDistributionList() } catch { $dl = $null }
          if ($dl -and $dl.PrimarySmtpAddress) { Write-Output $dl.PrimarySmtpAddress; exit }
          if ($r.AddressEntry.Address -match '@') { Write-Output $r.AddressEntry.Address; exit }
        }
      } catch { }
    }

    # 4) 주소를 못 구했으면 스토어 표시 이름이라도 알려준다 (목록의 이름-매칭용)
    if ($store.DisplayName) { Write-Output ('NAME=' + $store.DisplayName); exit }
  }
} catch { }
`;

function getActiveOutlookAccount() {
  return new Promise(resolve => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', PS_SCRIPT],
      { timeout: 30000, windowsHide: true },   // GAL 조회가 느린 PC가 있어 넉넉히
      (err, stdout) => {
        if (err) return resolve(null);
        const out = (stdout || '').trim().split(/\r?\n/)[0] || '';
        // 주소를 못 구한 스토어는 'NAME=<표시 이름>' 으로 온다 — 이름을 그대로 돌려준다
        if (out.startsWith('NAME=')) {
          const nm = out.slice(5).trim().toLowerCase();
          return resolve(nm || null);
        }
        resolve(/^[^\s@]+@[^\s@]+$/.test(out) ? out.toLowerCase() : null);
      });
  });
}

// Outlook에 붙어 있는 모든 사서함(스토어) 목록 — [{smtp, name}]
// 주소를 못 구해도 Exchange 사서함이면 이름으로 목록에 넣는다 (예: "Timeless Work" 공유 사서함).
// 주소도 없고 Exchange도 아닌 로컬 파일(PST)만 제외.
const PS_LIST = `
function Get-OutlookApp {
  $app = $null
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $app = $null }
  if ($null -eq $app) {
    # ROT 등록이 안 된 경우(권한 수준 차이 등)에도 붙을 수 있게 COM 개체로 직접 연결
    try { $app = New-Object -ComObject Outlook.Application } catch { $app = $null }
  }
  return $app
}
try {
  $ol = Get-OutlookApp
  if ($null -eq $ol) { throw 'NO_OUTLOOK' }
  $ses = $ol.Session
  $acct = @{}
  foreach ($a in $ses.Accounts) {
    try { if ($a.DeliveryStore) { $acct[$a.DeliveryStore.StoreID] = $a.SmtpAddress } } catch { }
  }
  foreach ($st in $ses.Stores) {
    $smtp = ''
    try { if ($acct.ContainsKey($st.StoreID)) { $smtp = $acct[$st.StoreID] } } catch { }
    if (-not $smtp) { try { if ($st.DisplayName -match '@') { $smtp = $st.DisplayName } } catch { } }
    if (-not $smtp) {
      $names = @()
      try { $names += $st.DisplayName } catch { }
      try { $names += $st.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x661B001F') } catch { }
      foreach ($n in $names) {
        if ((-not $n) -or $smtp) { continue }
        try {
          $r = $ses.CreateRecipient($n)
          $null = $r.Resolve()
          if ($r.Resolved) {
            try { $eu = $r.AddressEntry.GetExchangeUser() } catch { $eu = $null }
            if ($eu -and $eu.PrimarySmtpAddress) { $smtp = $eu.PrimarySmtpAddress }
            else {
              try { $dl = $r.AddressEntry.GetExchangeDistributionList() } catch { $dl = $null }
              if ($dl -and $dl.PrimarySmtpAddress) { $smtp = $dl.PrimarySmtpAddress }
              elseif ($r.AddressEntry.Address -match '@') { $smtp = $r.AddressEntry.Address }
            }
          }
        } catch { }
      }
    }
    $nm = ''
    try { $nm = $st.DisplayName } catch { }
    # Online Archive(온라인 보관) 스토어는 목록에서 제외 — 요약 대상 아님
    if ($nm -match '^\s*(Online Archive|In-Place Archive|온라인 보관)') { continue }
    if (-not $smtp) {
      # 주소를 못 구한 스토어: Exchange 사서함(공유 등)이면 이름으로라도 목록에 넣고, PST 등 로컬 파일만 뺀다
      $typ = 3
      try { $typ = [int]$st.ExchangeStoreType } catch { $typ = 3 }
      if ($typ -ge 3 -or -not $nm) { continue }
    }
    Write-Output ('MBOX=' + $smtp + [char]9 + $nm)
  }
} catch { }
`;

function listOutlookMailboxes() {
  return new Promise(resolve => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', PS_LIST],
      { timeout: 25000, windowsHide: true },
      (err, stdout) => {
        if (err) return resolve([]);
        const out = [], seen = new Set();
        for (const line of String(stdout || '').split(/\r?\n/)) {
          const m = /^MBOX=([^\t]*)\t?(.*)$/.exec(line.trim());
          if (!m) continue;
          let smtp = m[1].trim().toLowerCase();
          if (!/^[^\s@]+@[^\s@]+$/.test(smtp)) smtp = '';   // 주소 없는 사서함은 이름으로만
          const name = (m[2] || '').trim();
          if (!smtp && !name) continue;
          if (/^(online archive|in-place archive|온라인 보관)/i.test(name)) continue;   // 보관함 제외 (2중 안전망)
          const key = smtp || ('name:' + name.toLowerCase());
          if (seen.has(key)) continue;
          seen.add(key);
          out.push({ smtp, name });
        }
        resolve(out);
      });
  });
}

// 이 PC에서 어떤 Outlook이 실행 중인지 — classic(자동화 가능) / new(불가) / none
function outlookState() {
  return new Promise(resolve => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command',
        "$c = @(Get-Process -Name outlook -ErrorAction SilentlyContinue).Count; " +
        "$n = @(Get-Process -Name olk -ErrorAction SilentlyContinue).Count; " +
        "if ($c -gt 0) { 'classic' } elseif ($n -gt 0) { 'new' } else { 'none' }"],
      { timeout: 20000, windowsHide: true },
      (err, stdout) => {
        const v = String(stdout || '').trim().toLowerCase();
        resolve(['classic', 'new', 'none'].includes(v) ? v : 'none');
      });
  });
}

module.exports = { getActiveOutlookAccount, listOutlookMailboxes, outlookState };
