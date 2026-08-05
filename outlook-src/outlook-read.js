// 실행 중인 Outlook에서 직접 읽기 (COM) — 서버 권한·로그인·테넌트와 무관하게
// "그 PC의 Outlook에 보이는 사서함"이면 무엇이든 읽는다 (teamforce 같은 다른 회사 계정 포함).
// Outlook이 꺼져 있으면 동작하지 않는다.
const { execFile } = require('child_process');

// 스토어(사서함)의 SMTP 주소를 구하는 공통 PowerShell 조각 — outlook-detect.js와 같은 규칙
const PS_RESOLVE = `
function Get-OutlookApp {
  $app = $null
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $app = $null }
  if ($null -eq $app) {
    # ROT 등록이 안 된 경우(권한 수준 차이 등)에도 붙을 수 있게 COM 개체로 직접 연결
    try { $app = New-Object -ComObject Outlook.Application } catch { $app = $null }
  }
  return $app
}
function Get-StoreSmtp($ses, $st, $acct) {
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
  return $smtp
}
# 사서함 지정값(주소 또는 스토어 표시 이름)으로 스토어 찾기 — 둘 다 대소문자 무시
function Find-Store($ses, $acct, $want) {
  foreach ($st in $ses.Stores) {
    $smtp = Get-StoreSmtp $ses $st $acct
    $dn = ''
    try { $dn = '' + $st.DisplayName } catch { }
    # 기본 사서함 자동 선택에서 Online Archive(온라인 보관)는 건너뛴다 (이름으로 콕 집으면 허용)
    if ($dn -match '^\s*(Online Archive|In-Place Archive|온라인 보관)') {
      if ($want -eq '' -or -not ($dn -and $dn.ToLower() -eq $want)) { continue }
    }
    if ($want -eq '') {
      if (-not $smtp) { continue }
      return @{ st = $st; box = $smtp }
    }
    if (($smtp -and $smtp.ToLower() -eq $want) -or ($dn -and $dn.ToLower() -eq $want)) {
      $box = $smtp
      if (-not $box) { $box = $dn }
      return @{ st = $st; box = $box }
    }
  }
  return $null
}
`;

const PS_READ = PS_RESOLVE + `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$want = ('' + $env:OBAR_MBOX).Trim().ToLower()
$max = 0; try { $max = [int]$env:OBAR_MAX } catch { }
if ($max -le 0) { $max = 200 }
$blen = 0; try { $blen = [int]$env:OBAR_BODY } catch { }
if ($blen -le 0) { $blen = 6000 }
$res = [pscustomobject]@{ mailbox = ''; items = @(); error = ''; detail = ''; how = ''; inboxCount = 0; flagCount = 0 }
try {
  $ol = Get-OutlookApp
  if ($null -eq $ol) { throw 'NO_OUTLOOK' }
  $ses = $ol.Session
  $acct = @{}
  foreach ($a in $ses.Accounts) {
    try { if ($a.DeliveryStore) { $acct[$a.DeliveryStore.StoreID] = $a.SmtpAddress } } catch { }
  }
  $found = Find-Store $ses $acct $want
  if ($null -eq $found) {
    $res.error = 'MAILBOX_NOT_FOUND'
  } else {
    $target = $found.st
    $res.mailbox = $found.box
    $inbox = $target.GetDefaultFolder(6)

    # 읽을 폴더 목록 — OBAR_FOLDERS 가 없으면 Inbox 만 (기존 동작).
    # 형식: 폴더들은 [char]30 으로 구분, 한 폴더의 경로 단계는 [char]31 로 구분, '.' = Inbox 자체
    $folders = New-Object System.Collections.ArrayList
    $specsRaw = '' + $env:OBAR_FOLDERS
    if (-not $specsRaw) {
      $null = $folders.Add(@{ f = $inbox; label = 'Inbox' })
    } else {
      foreach ($spec in $specsRaw.Split([char]30)) {
        if (-not $spec) { continue }
        if ($spec -eq '.') { $null = $folders.Add(@{ f = $inbox; label = 'Inbox' }); continue }
        $f = $inbox
        $ok = $true
        $lab = ''
        foreach ($seg in $spec.Split([char]31)) {
          if (-not $seg) { continue }
          try { $f = $f.Folders.Item($seg) } catch { $f = $null }
          if ($null -eq $f) { $ok = $false; break }
          if ($lab) { $lab = $lab + '/' + $seg } else { $lab = $seg }
        }
        if ($ok -and $f) { $null = $folders.Add(@{ f = $f; label = $lab }) }
      }
      if ($folders.Count -eq 0) { $null = $folders.Add(@{ f = $inbox; label = 'Inbox' }) }
    }

    $list = New-Object System.Collections.ArrayList
    $n = 0
    $hows = New-Object System.Collections.ArrayList
    foreach ($fol in $folders) {
      if ($n -ge $max) { break }
      $all = $fol.f.Items
      $how = ''
      $flag = $null
      # 방법 1) Outlook 필드 이름으로 걸러내기
      try {
        $c = $all.Restrict('[FlagStatus] = 2')
        if ($c.Count -gt 0) { $flag = $c; $how = 'FlagStatus' }
      } catch { }
      # 방법 2) MAPI 속성(PR_FLAG_STATUS)으로 걸러내기
      if ($null -eq $flag) {
        try {
          $c = $all.Restrict('@SQL="http://schemas.microsoft.com/mapi/proptag/0x10900003" = 2')
          if ($c.Count -gt 0) { $flag = $c; $how = 'proptag' }
        } catch { }
      }
      # 방법 3) 걸러내기가 안 되는 사서함(계정 종류에 따라 다름) — 최근 것부터 직접 훑는다
      if ($null -eq $flag) {
        try { $all.Sort('[ReceivedTime]', $true) } catch { }
        $picked = New-Object System.Collections.ArrayList
        $seen = 0
        foreach ($x in $all) {
          $seen++
          if ($seen -gt 3000 -or $picked.Count -ge $max) { break }
          try {
            $fs = 0
            try { $fs = [int]$x.FlagStatus } catch { $fs = 0 }
            if ($fs -ne 2) {
              try { if ($x.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x10900003') -eq 2) { $fs = 2 } } catch { }
            }
            if ($fs -eq 2) { $null = $picked.Add($x) }
          } catch { }
        }
        $flag = $picked
        $how = 'scan'
      }
      if ($hows -notcontains $how) { $null = $hows.Add($how) }
      try { $res.inboxCount = $res.inboxCount + $all.Count } catch { }
      try { $res.flagCount = $res.flagCount + $flag.Count } catch { }
      foreach ($it in $flag) {
        if ($n -ge $max) { break }
        try {
          if ($it.Class -eq 40 -or $it.Class -eq 69) { continue }   # 연락처·작업 등은 제외
          $addr = ''
          try { $addr = $it.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x5D01001F') } catch { }
          if (-not $addr) { try { $addr = $it.SenderEmailAddress } catch { } }
          $txt = ''
          try { $txt = '' + $it.Body } catch { }
          if ($txt.Length -gt $blen) { $txt = $txt.Substring(0, $blen) }
          $rt = ''
          try { $rt = $it.ReceivedTime.ToString('o') } catch { }
          $null = $list.Add([pscustomobject]@{
            id = $it.EntryID
            subject = ('' + $it.Subject)
            name = ('' + $it.SenderName)
            addr = ('' + $addr)
            received = $rt
            text = $txt
            folder = ('' + $fol.label)
          })
          $n++
        } catch { }
      }
    }
    $res.how = ($hows -join '+')
    $res.items = $list.ToArray()
  }
} catch {
  $res.error = 'OUTLOOK_UNAVAILABLE'
  try { $res.detail = ('' + $_.Exception.Message) } catch { }
}
ConvertTo-Json -InputObject $res -Depth 4 -Compress
`;

function runPs(script, env, timeout = 180000) {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script],
      { timeout, windowsHide: true, maxBuffer: 128 * 1024 * 1024, encoding: 'utf8', env: { ...process.env, ...env } },
      (err, stdout) => {
        if (err && !stdout) return reject(err);
        resolve(String(stdout || ''));
      });
  });
}

// 대상 사서함의 Inbox(+선택한 하위 폴더)에서 flag된 메일을 읽어 Graph와 같은 모양으로 돌려준다
// mailbox 는 주소 또는 스토어 표시 이름("Timeless Work") 모두 허용. 비어 있으면 Outlook의 기본 사서함.
// folders: ['.','sub','sub하위'] 형태의 spec 배열 ('.'=Inbox). 비어 있으면 Inbox 만.
async function readFlagged(mailbox, { max = 200, bodyChars = 6000, folders = [] } = {}) {
  let raw;
  try {
    raw = await runPs(PS_READ, {
      OBAR_MBOX: mailbox || '',
      OBAR_MAX: String(max),
      OBAR_BODY: String(bodyChars),
      OBAR_FOLDERS: (folders || []).join('\u001e'),
    });
  } catch (e) {
    throw new Error('Outlook을 읽지 못했습니다. Outlook이 켜져 있는지 확인해 주세요. (' + e.message + ')');
  }
  let data;
  try { data = JSON.parse(raw.trim() || '{}'); } catch (e) {
    throw new Error('Outlook 응답을 해석하지 못했습니다. Outlook이 켜져 있는지 확인해 주세요.');
  }
  if (data.error === 'OUTLOOK_UNAVAILABLE') {
    throw new Error('Outlook에 연결하지 못했습니다.' + (data.detail ? ' (' + data.detail + ')' : ''));
  }
  if (data.error === 'MAILBOX_NOT_FOUND') {
    throw new Error(`"${mailbox}" 사서함을 Outlook에서 찾지 못했습니다. 목록에서 다시 골라주세요.`);
  }
  const items = Array.isArray(data.items) ? data.items : (data.items ? [data.items] : []);
  const msgs = items.map(it => ({
    id: it.id,
    subject: it.subject || '',
    from: { emailAddress: { name: it.name || '', address: it.addr || '' } },
    receivedDateTime: it.received || new Date(0).toISOString(),
    bodyPreview: String(it.text || '').replace(/\s+/g, ' ').slice(0, 200),
    _text: String(it.text || ''),
    _folder: it.folder || 'Inbox',
  }));
  return {
    mailbox: (data.mailbox || mailbox || '').toLowerCase(),
    msgs,
    how: data.how || '',
    inboxCount: data.inboxCount || 0,
    flagCount: data.flagCount || 0,
  };
}

// flag 해제 (리포트의 Unflag 버튼) — EntryID로 항목을 찾아 깃발을 끈다
const PS_UNFLAG = `
function Get-OutlookApp {
  $app = $null
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $app = $null }
  if ($null -eq $app) {
    # ROT 등록이 안 된 경우(권한 수준 차이 등)에도 붙을 수 있게 COM 개체로 직접 연결
    try { $app = New-Object -ComObject Outlook.Application } catch { $app = $null }
  }
  return $app
}
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
  $ol = Get-OutlookApp
  if ($null -eq $ol) { throw 'NO_OUTLOOK' }
  $it = $ol.Session.GetItemFromID($env:OBAR_ID)
  if ($null -eq $it) { Write-Output 'GONE'; exit }
  try { $it.ClearTaskFlag() } catch { try { $it.FlagStatus = 0 } catch { } }
  try { $it.Save() } catch { }
  Write-Output 'OK'
} catch { Write-Output 'ERR' }
`;

async function unflagLocal(entryId) {
  const out = (await runPs(PS_UNFLAG, { OBAR_ID: entryId }, 60000)).trim();
  if (out.includes('OK')) return 'unflagged';
  if (out.includes('GONE')) return 'gone';
  throw new Error('Outlook에서 flag를 해제하지 못했습니다.');
}

// 리포트에서 제목을 클릭하면 Outlook에서 그 메일을 연다
const PS_OPEN = `
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
  $it = $ol.Session.GetItemFromID($env:OBAR_ID)
  if ($null -ne $it) { $it.Display() }
  Write-Output 'OK'
} catch { Write-Output 'ERR' }
`;

async function openLocal(entryId) {
  const out = (await runPs(PS_OPEN, { OBAR_ID: entryId }, 60000)).trim();
  if (out.includes('OK')) return 'opened';
  throw new Error('Outlook에서 메일을 열지 못했습니다.');
}

// 대상 사서함의 Inbox 하위 폴더 목록 (바의 폴더 선택 창용)
// 출력 각 항목: { spec, display } — spec 은 readFlagged 의 folders 에 그대로 넘기는 값
const PS_FOLDERS = PS_RESOLVE + `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$want = ('' + $env:OBAR_MBOX).Trim().ToLower()
try {
  $ol = Get-OutlookApp
  if ($null -eq $ol) { throw 'NO_OUTLOOK' }
  $ses = $ol.Session
  $acct = @{}
  foreach ($a in $ses.Accounts) {
    try { if ($a.DeliveryStore) { $acct[$a.DeliveryStore.StoreID] = $a.SmtpAddress } } catch { }
  }
  $found = Find-Store $ses $acct $want
  if ($null -eq $found) { Write-Output 'ERR=MAILBOX_NOT_FOUND'; exit }
  $inbox = $found.st.GetDefaultFolder(6)
  Write-Output ('FLD=.' + [char]9 + 'Inbox')
  # 폴더 하나가 오류를 내도(권한·동기화 문제) 전체 목록이 죽지 않게 폴더별로 감싼다
  function Walk($fld, $spec, $disp, $depth) {
    $kids = @()
    try { $kids = @($fld.Folders) } catch { return }
    foreach ($c in $kids) {
      try {
        $nm = ''
        try { $nm = '' + $c.Name } catch { }
        if (-not $nm) { continue }
        $ns = $nm
        if ($spec -ne '.') { $ns = $spec + [char]31 + $nm }
        $nd = $disp + ' / ' + $nm
        Write-Output ('FLD=' + $ns + [char]9 + $nd)
        if ($depth -lt 3) { Walk $c $ns $nd ($depth + 1) }
      } catch { }
    }
  }
  Walk $inbox '.' 'Inbox' 1
} catch {
  $d = ''
  try { $d = ' ' + $_.Exception.Message } catch { }
  Write-Output ('ERR=OUTLOOK_UNAVAILABLE' + $d)
}
`;

// 최근 N시간 Inbox 메일 + 그 기간의 보낸 메일 대화ID 목록 (Review Daily용)
// verb(PR_LAST_VERB_EXECUTED): 102=Reply, 103=ReplyAll, 104=Forward
const PS_RECENT = PS_RESOLVE + `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$want = ('' + $env:OBAR_MBOX).Trim().ToLower()
$hrs = 0; try { $hrs = [double]$env:OBAR_HOURS } catch { }
if ($hrs -le 0) { $hrs = 24 }
$max = 0; try { $max = [int]$env:OBAR_MAX } catch { }
if ($max -le 0) { $max = 400 }
$blen = 0; try { $blen = [int]$env:OBAR_BODY } catch { }
if ($blen -le 0) { $blen = 700 }
$res = [pscustomobject]@{ mailbox = ''; items = @(); sent = @(); error = ''; detail = ''; scanned = 0 }
try {
  $ol = Get-OutlookApp
  if ($null -eq $ol) { throw 'NO_OUTLOOK' }
  $ses = $ol.Session
  $acct = @{}
  foreach ($a in $ses.Accounts) {
    try { if ($a.DeliveryStore) { $acct[$a.DeliveryStore.StoreID] = $a.SmtpAddress } } catch { }
  }
  $found = Find-Store $ses $acct $want
  if ($null -eq $found) {
    $res.error = 'MAILBOX_NOT_FOUND'
  } else {
    $res.mailbox = $found.box
    $cut = (Get-Date).AddHours(-1 * $hrs)
    $inbox = $found.st.GetDefaultFolder(6)
    $all = $inbox.Items
    try { $all.Sort('[ReceivedTime]', $true) } catch { }
    $list = New-Object System.Collections.ArrayList
    $n = 0
    $seen = 0
    foreach ($x in $all) {
      $seen++
      if ($seen -gt 2000 -or $n -ge $max) { break }
      $rt = $null
      try { $rt = $x.ReceivedTime } catch { }
      if ($null -eq $rt) { continue }
      if ($rt -lt $cut) { break }
      $cls = 0
      try { $cls = [int]$x.Class } catch { }
      if ($cls -ne 43) { continue }
      $verb = 0
      try { $verb = [int]$x.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x10810003') } catch { }
      $cid = ''
      try { $cid = '' + $x.ConversationID } catch { }
      $addr = ''
      try { $addr = $x.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x5D01001F') } catch { }
      if (-not $addr) { try { $addr = $x.SenderEmailAddress } catch { } }
      $ur = $false
      try { $ur = [bool]$x.UnRead } catch { }
      $txt = ''
      try { $txt = '' + $x.Body } catch { }
      if ($txt.Length -gt $blen) { $txt = $txt.Substring(0, $blen) }
      $null = $list.Add([pscustomobject]@{
        id = $x.EntryID
        subject = ('' + $x.Subject)
        name = ('' + $x.SenderName)
        addr = ('' + $addr)
        received = $rt.ToString('o')
        preview = $txt
        verb = $verb
        conv = $cid
        unread = $ur
      })
      $n++
    }
    $res.scanned = $seen
    $res.items = $list.ToArray()
    # 같은 기간의 '보낸 메일' 대화 ID — 대상 사서함 + 내 기본 사서함 둘 다 (공유 사서함 답장이 내 보낸함에 저장되는 경우)
    $sc = New-Object System.Collections.ArrayList
    $stores = New-Object System.Collections.ArrayList
    $null = $stores.Add($found.st)
    try { $ds = $ses.DefaultStore; if ($ds -and $ds.StoreID -ne $found.st.StoreID) { $null = $stores.Add($ds) } } catch { }
    foreach ($s2 in $stores) {
      $sf = $null
      try { $sf = $s2.GetDefaultFolder(5) } catch { }
      if ($null -eq $sf) { continue }
      $si = $sf.Items
      try { $si.Sort('[SentOn]', $true) } catch { }
      $c2 = 0
      foreach ($y in $si) {
        $c2++
        if ($c2 -gt 500) { break }
        $so = $null
        try { $so = $y.SentOn } catch { }
        if ($null -eq $so) { continue }
        if ($so -lt $cut) { break }
        $cid2 = ''
        try { $cid2 = '' + $y.ConversationID } catch { }
        if ($cid2) { $null = $sc.Add($cid2) }
      }
    }
    $res.sent = $sc.ToArray()
  }
} catch {
  $res.error = 'OUTLOOK_UNAVAILABLE'
  try { $res.detail = ('' + $_.Exception.Message) } catch { }
}
ConvertTo-Json -InputObject $res -Depth 4 -Compress
`;

// 최근 hours시간 Inbox 메일 + 보낸-대화 세트 (Review Daily)
async function readRecent(mailbox, { hours = 24, bodyChars = 700, max = 400 } = {}) {
  let raw;
  try {
    raw = await runPs(PS_RECENT, {
      OBAR_MBOX: mailbox || '',
      OBAR_HOURS: String(hours),
      OBAR_BODY: String(bodyChars),
      OBAR_MAX: String(max),
    });
  } catch (e) {
    throw new Error('Outlook을 읽지 못했습니다. Outlook이 켜져 있는지 확인해 주세요. (' + e.message + ')');
  }
  let data;
  try { data = JSON.parse(raw.trim() || '{}'); } catch (e) {
    throw new Error('Outlook 응답을 해석하지 못했습니다. Outlook이 켜져 있는지 확인해 주세요.');
  }
  if (data.error === 'OUTLOOK_UNAVAILABLE') {
    throw new Error('Outlook에 연결하지 못했습니다.' + (data.detail ? ' (' + data.detail + ')' : ''));
  }
  if (data.error === 'MAILBOX_NOT_FOUND') {
    throw new Error(`"${mailbox}" 사서함을 Outlook에서 찾지 못했습니다. 목록에서 다시 골라주세요.`);
  }
  const items = Array.isArray(data.items) ? data.items : (data.items ? [data.items] : []);
  const sent = Array.isArray(data.sent) ? data.sent : (data.sent ? [data.sent] : []);
  return {
    mailbox: (data.mailbox || mailbox || '').toLowerCase(),
    items,
    sentSet: new Set(sent.filter(Boolean)),
    scanned: data.scanned || 0,
  };
}

async function listFolders(mailbox) {
  // 온라인 모드 공유 사서함은 폴더 순회가 느릴 수 있어 넉넉히 기다린다
  const raw = await runPs(PS_FOLDERS, { OBAR_MBOX: mailbox || '' }, 120000);
  const out = [];
  for (const line of String(raw || '').split(/\r?\n/)) {
    const t = line.replace(/\r$/, '');
    if (t.startsWith('ERR=')) throw new Error(t.slice(4));
    if (!t.startsWith('FLD=')) continue;
    const i = t.indexOf('\t');
    if (i < 0) continue;
    const spec = t.slice(4, i);
    const display = t.slice(i + 1).trim();
    if (spec) out.push({ spec, display });
  }
  return out;
}

module.exports = { readFlagged, readRecent, unflagLocal, openLocal, listFolders };
