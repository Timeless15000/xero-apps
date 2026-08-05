# auto-publish.ps1 — 바뀐 파일을 자동으로 GitHub 에 올린다 (Brian PC 전용)
# XERO 바 / OUTLOOK 바가 10분마다 조용히 실행한다. 사람이 누를 것은 없다.
# 안전장치: ①동시 실행 방지 ②방금 저장된 파일이 있으면 다음 차례로 미룸(작업 중간 업로드 방지)
$ErrorActionPreference = 'SilentlyContinue'

$here = $PSScriptRoot
$parent = Split-Path $here -Parent
$repos = @($here, (Join-Path $parent 'Outlook'), (Join-Path $parent 'xero-apps')) | Select-Object -Unique

$log = Join-Path $here 'auto-publish.log'
function Say($m) {
    $line = '[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ' + $m
    Add-Content -Path $log -Value $line -Encoding UTF8
}

# 로그가 너무 커지면 줄인다
if ((Test-Path $log) -and ((Get-Item $log).Length -gt 300000)) {
    $keep = Get-Content $log -Tail 300
    Set-Content -Path $log -Value $keep -Encoding UTF8
}

# ── 동시 실행 방지 (두 바가 동시에 불러도 하나만 돈다)
$lock = Join-Path $env:TEMP 'xero_autopublish.lock'
if (Test-Path $lock) {
    $mins = (New-TimeSpan -Start (Get-Item $lock).LastWriteTime -End (Get-Date)).TotalMinutes
    if ($mins -lt 5) { exit }
}
Set-Content -Path $lock -Value (Get-Date).ToString()

# ── git 찾기 (PATH → GitHub Desktop 내장)
$git = ''
$c = Get-Command git -ErrorAction SilentlyContinue
if ($c) { $git = $c.Source }
if (-not $git) {
    $d = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'GitHubDesktop') -Directory -Filter 'app-*' -ErrorAction SilentlyContinue |
         Sort-Object Name -Descending | Select-Object -First 1
    if ($d) { $git = Join-Path $d.FullName 'resources\app\git\cmd\git.exe' }
}
if ((-not $git) -or (-not (Test-Path $git))) { Say 'git 을 찾지 못했습니다 (GitHub Desktop 설치 확인)'; exit }

foreach ($r in $repos) {
    if (-not (Test-Path (Join-Path $r '.git'))) { continue }
    $name = Split-Path $r -Leaf
    Push-Location $r

    # OneDrive 가 .git 을 잠그지 않도록 자동 정리 기능을 끈다 (Publish bat 과 같은 설정)
    & $git config gc.auto 0 | Out-Null
    & $git config gc.autoDetach false | Out-Null
    & $git config maintenance.auto false | Out-Null
    & $git config fetch.writeCommitGraph false | Out-Null
    $gclog = Join-Path $r '.git\gc.log'
    if (Test-Path $gclog) { Remove-Item $gclog -Force -ErrorAction SilentlyContinue }

    $st = & $git status --porcelain 2>$null
    if (-not $st) { Pop-Location; continue }

    # 방금 수정된 파일이 있으면 이번엔 건너뛴다 (저장 중간에 올리지 않으려고)
    $tooNew = $false
    $cut = (Get-Date).AddMinutes(-2)
    foreach ($line in $st) {
        $rel = ($line.Substring(3)).Trim('"')
        if ($rel -match ' -> ') { $rel = ($rel -split ' -> ')[-1] }
        $full = Join-Path $r $rel
        if (Test-Path $full -PathType Leaf) {
            if ((Get-Item $full).LastWriteTime -gt $cut) { $tooNew = $true; break }
        }
    }
    if ($tooNew) { Say ($name + ' : 방금 저장된 파일이 있어 다음 차례로 미룸'); Pop-Location; continue }

    $n = ($st | Measure-Object).Count
    & $git add -A 2>$null | Out-Null
    & $git commit -m ('auto publish ' + (Get-Date -Format 'yyyy-MM-dd HH:mm')) 2>$null | Out-Null
    & $git pull --no-edit 2>$null | Out-Null
    $out = & $git push 2>&1
    $txt = ($out | Out-String).Trim()
    if ($LASTEXITCODE -eq 0) {
        Say ($name + ' : 파일 ' + $n + '개 업로드 완료')
    } else {
        Say ($name + ' : 업로드 실패 - ' + $txt)
    }
    Pop-Location
}
