#Requires AutoHotkey v2.0
#SingleInstance Force

; ================= OUTLOOK Desktop Bar =================
; XERO 바와 같은 조작감의 데스크톱 바. 차이점: 버튼이 단축키를 보내는 대신
; 이 저장소의 Node 명령을 직접 실행한다 (Tampermonkey 불필요).
; - Flagged Summary: flag된 메일을 AI 요약과 함께 HTML 리포트로 정리해 브라우저로 연다.
;   메일은 실행 중인 Outlook 에서 직접 읽는다 → 회사 계정 구분·서버 권한·로그인 불필요.
;   지금 Outlook 화면에 보고 있는 사서함(공유·추가 사서함 포함)을 자동으로 대상으로 삼는다.
;   특정 사서함을 항상 보려면 tools 에 mbox 를 지정한 버튼을 두면 된다.
; EDIT: 보여줄 버튼만 체크 → SAVE / X(취소). 크기 조절: 창 오른쪽 아래 코너 드래그. 설정은 저장됨.

ini := A_ScriptDir "\OUTLOOK_bar.ini"
UPDATE_URL := "https://raw.githubusercontent.com/Timeless15000/xero-apps/main/OUTLOOK_bar.ahk"
; 프로그램 파일(src)도 바와 똑같이 GitHub 에서 직접 받는다 (회사 폴더 복사에만 의존하지 않도록).
; config.json(API 키)·tone-guide.md·서명 파일은 일부러 올리지 않는다 → 그것들은 계속 회사 폴더에서 온다.
SRC_BASE := "https://raw.githubusercontent.com/Timeless15000/xero-apps/main/outlook-src/"
SRC_FILES := ["ai.js", "auth.js", "flagged.js", "graph.js", "index.js", "outlook-detect.js"
            , "outlook-read.js", "review-daily.js", "srcver.js", "unflag-server.js", "package.json"]
AUTO_UPDATE := !InStr(A_ScriptDir, "GitHub")   ; 관리자 원본 폴더에서는 자동 업데이트 안 함

APPVER := 27                              ; 앱 버전 — 이 폴더의 무엇이든 고치면 +1 (바 파일뿐 아니라 src\*.js 포함)
DATEVER := "02/08/2026"                 ; 오프라인 기본값. 아래에서 파일 수정날짜로 자동 대체.
try DATEVER := FormatTime(FileGetTime(A_ScriptFullPath, "M"), "dd/MM/yyyy")  ; 이 파일 마지막 수정일 = 버전 날짜

; 공유 사서함 버튼을 늘리려면 {label, c, id:"영문id", mbox:"<메일주소>"} 한 줄을 추가하면 된다
tools := [
    {label:"Flagged Summary", c:"0F6CBD", id:"flaggedsummary", mbox:""},   ; 보고 있는 사서함 자동
    {label:"Review Daily",    c:"C8511B", id:"reviewdaily",    mbox:""}    ; 최근 24시간 미답장 메일
]

enabled := Map()
for t in tools
    enabled[t.id] := (IniRead(ini, "tools", t.id, "1") = "1")

editMode := false
busy := false
posX := 20
posY := 420
scale := IniRead(ini, "cfg", "scale", "1") + 0
if (scale < 0.6)
    scale := 0.6
if (scale > 3)
    scale := 3
opacity := IniRead(ini, "cfg", "opacity", "100") + 0
if (opacity < 20)
    opacity := 20
if (opacity > 100)
    opacity := 100
g := ""
fg := ""
checks := Map()
pickBox := ""
pickMap := Map()
items := []
srcVer := ""        ; 프로그램 파일(src) 버전 — APPVER 와 달라지면 경고
srcWarned := false
tipG := ""          ; 로딩 상태창 (바 밑에 붙어 같이 움직인다)
stickyMsg := ""     ; 작업이 끝날 때까지 유지할 메시지
tipSeq := 0

; 첫 실행 시 바탕화면에 "OUTLOOK Bar" 아이콘 자동 생성
try {
    _lnk := A_Desktop "\OUTLOOK Bar.lnk"
    if !FileExist(_lnk)
        FileCreateShortcut(A_ScriptFullPath, _lnk, A_ScriptDir)
}

; 예전에 켜둔 바가 있으면 닫기 (중복 방지) — 옛 제목("OUTLOOK Bar")과 새 제목("Outlook (") 둘 다
SetTitleMatchMode(2)
try {
    for _pat in ["Outlook (", "OUTLOOK Bar"]
        for _hw in WinGetList(_pat)
            if (_hw != 0)
                try WinClose("ahk_id " _hw)
}

Build()

; ===== 자동 업로드 (Brian PC 전용) =====
; 이 바가 관리자 원본 폴더에서 돌 때만, 바뀐 파일을 10분마다 조용히 GitHub 에 올린다.
; 직원 PC 사본은 GitHub 폴더가 아니므로 절대 실행되지 않는다. 사람이 누를 것은 없다.
AUTO_PUBLISH := InStr(A_ScriptDir, "GitHub")
if AUTO_PUBLISH {
    SetTimer(AutoPublish, -90000)              ; 켠 뒤 1분 30초에 한 번
    SetTimer(AutoPublish, 10 * 60 * 1000)      ; 이후 10분마다
}

AutoPublish() {
    ps := A_ScriptDir "\auto-publish.ps1"
    if !FileExist(ps)
        ps := A_ScriptDir "\..\xero-apps\auto-publish.ps1"
    if !FileExist(ps)
        return
    try Run('powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' ps '"', , "Hide")
}

if AUTO_UPDATE {
    SetTimer(() => CheckUpdate(true), -4000)              ; 켠 뒤 4초 후 1회
    SetTimer(() => CheckUpdate(true), 2 * 60 * 60 * 1000) ; 이후 2시간마다
    ; 버전 대조는 CheckUpdate 안에서 RefreshProgram 이 끝난 직후에 한다 (경합 방지)
} else {
    SetTimer(() => CheckSrcVersion(true), -1500)
}

Build() {
    global g, tools, enabled, editMode, posX, posY, checks, scale, items, opacity, APPVER, DATEVER
    if IsObject(g) {
        try {
            WinGetPos(&px, &py, , , "ahk_id " g.Hwnd)
            if (px != "") {
                posX := px
                posY := py
            }
        }
        g.Destroy()
    }
    checks := Map()
    items := []
    g := Gui("+AlwaysOnTop +Resize +ToolWindow -MaximizeBox -MinimizeBox", "Outlook (" DATEVER ") v" APPVER)
    g.BackColor := "16324A"
    g.OnEvent("Close", (*) => ExitApp())
    g.OnEvent("Size", GuiResize)

    M := 8
    gap := 6
    CW := 180
    y := M
    sld := g.Add("Slider", "x" M " y" y " w" CW " h20 Range20-100 Line1 Page10 ToolTip")
    sld.Value := opacity
    sld.OnEvent("Change", (ctrl, *) => SetOpacity(ctrl.Value))
    AddItem(sld, M, y, CW, 20)
    y += 20 + gap
    if editMode {
        sv := g.Add("Text", "x" M " y" y " w144 h24 Center 0x200 Background1E88E5 cWhite", "SAVE")
        sv.OnEvent("Click", (*) => SaveEdit())
        AddItem(sv, M, y, 144, 24)
        cx := g.Add("Text", "x" (M + 144 + gap) " y" y " w" (CW - 144 - gap) " h24 Center 0x200 BackgroundB71C1C cWhite", "X")
        cx.OnEvent("Click", (*) => CancelEdit())
        AddItem(cx, M + 144 + gap, y, CW - 144 - gap, 24)
        y += 24 + gap
        for t in tools {
            cb := g.Add("CheckBox", "x" M " y" y " w" CW " h22 cWhite " (enabled[t.id] ? "Checked" : ""), t.label)
            checks[t.id] := cb
            AddItem(cb, M, y, CW, 22)
            y += 22 + gap
        }
    } else {
        et := g.Add("Text", "x" M " y" y " w" CW " h24 Center 0x200 Background455A64 cWhite", "EDIT")
        et.OnEvent("Click", (*) => EnterEdit())
        AddItem(et, M, y, CW, 24)
        y += 24 + gap
        for t in tools {
            if !enabled[t.id]
                continue
            b := g.Add("Text", "x" M " y" y " w" CW " h32 Center 0x200 Background" t.c " cWhite", t.label)
            b.OnEvent("Click", RunTool.Bind(t.id))
            AddItem(b, M, y, CW, 32)
            y += 32 + gap
        }
        ; UPDATE 버튼 — 항상 맨 아래 고정. 도구가 아니므로 EDIT 목록에 넣지 않는다
        ; (뭔가 이상할 때 직원이 바로 누를 것이 있어야 하므로 숨길 수 없게 한다)
        ub := g.Add("Text", "x" M " y" y " w" CW " h26 Center 0x200 Background546E7A cWhite", "UPDATE")
        ub.OnEvent("Click", UpdateNow)
        AddItem(ub, M, y, CW, 26)
        y += 26 + gap
    }
    baseW := M + CW + M
    baseH := (y - gap) + M
    g.Show("x" posX " y" posY " w" Round(baseW * scale) " h" Round(baseH * scale))
    Relayout()
    WinSetTransparent(Round(opacity * 2.55), "ahk_id " g.Hwnd)
}

AddItem(c, x, y, w, h) {
    global items
    items.Push({c: c, x: x, y: y, w: w, h: h})
}

Relayout() {
    global items, scale
    FS := Round(10 * scale)
    if FS < 6
        FS := 6
    for it in items {
        if (it.c.Type != "Slider")
            it.c.SetFont("s" FS " Bold", "Segoe UI")
        it.c.Move(Round(it.x * scale), Round(it.y * scale), Round(it.w * scale), Round(it.h * scale))
    }
}

GuiResize(thisGui, minMax, w, h) {
    global scale
    if (minMax = -1)
        return
    ns := w / 196.0
    if (ns < 0.6)
        ns := 0.6
    if (ns > 3)
        ns := 3
    scale := ns
    Relayout()
    SetTimer(SaveScale, -500)
}

SaveScale() {
    global scale, ini
    IniWrite(Round(scale, 3), ini, "cfg", "scale")
}

SetOpacity(v) {
    global g, opacity
    opacity := v
    WinSetTransparent(Round(v * 2.55), "ahk_id " g.Hwnd)
    SetTimer(SaveOpacity, -500)
}

SaveOpacity() {
    global opacity, ini
    IniWrite(opacity, ini, "cfg", "opacity")
}

EnterEdit() {
    global editMode
    editMode := true
    Build()
}

SaveEdit() {
    global editMode, enabled, checks, ini
    for id, cb in checks
        enabled[id] := cb.Value ? true : false
    for id, v in enabled
        IniWrite(v ? "1" : "0", ini, "tools", id)
    editMode := false
    Build()
}

CancelEdit() {
    global editMode
    editMode := false
    Build()
}

; ================= 버튼 동작 =================
RunTool(id, *) {
    global busy, tools
    for t in tools {
        if (t.id = id) {
            mb := t.HasOwnProp("mbox") ? t.mbox : ""
            if (t.id = "reviewdaily")
                RunReviewDaily(mb)
            else
                RunFlaggedSummary(mb)
            return
        }
    }
}

; 사서함 선택 공통 흐름 (목록 메뉴). 반환: Map("ok",0|1, "box","")
; ok=0 → 취소/실패(중단). ok=1 & box="" → 로그인한 본인 계정(기본 사서함)으로 진행.
PickMailboxFlow() {
    global pickBox, pickMap
    r := Map("ok", 0, "box", "")
    Tip("사서함 목록 확인 중...", 0)
    out := CaptureNode(' /c node "src\index.js" --list-mailboxes')
    HideTip()

    cur := ""
    if RegExMatch(out, "CUR=([^\r\n]*)", &mc)
        cur := Trim(mc[1], " `t")

    ; 주소를 못 구한 공유 사서함(예: "Timeless Work")은 smtp가 비어 있고 이름만 온다 — 이름으로 지정해 실행한다
    pickMap := Map()
    labels := []
    pos := 1
    while (hit := RegExMatch(out, "m)^MBOX=([^\t\r\n]*)\t?([^\r\n]*)$", &m, pos)) {
        pos := hit + StrLen(m[0]) + 1
        smtp := Trim(m[1], " `t")
        nm := Trim(m[2], " `t")
        if (smtp = "" && nm = "")
            continue
        lbl := (smtp = "") ? nm : ((nm != "" && nm != smtp) ? nm "  (" smtp ")" : smtp)
        if ((smtp != "" && smtp = cur) || (smtp = "" && nm = cur))
            lbl := lbl "   ← 지금 보는 폴더"
        if !pickMap.Has(lbl) {
            pickMap[lbl] := (smtp != "") ? smtp : nm
            labels.Push(lbl)
        }
    }

    if (labels.Length = 0) {
        state := ""
        if RegExMatch(out, "STATE=([^\r\n]*)", &ms)
            state := Trim(ms[1], " `t")
        if (state = "new") {
            MsgBox("이 PC는 새 Outlook(New Outlook) 을 쓰고 있어 메일을 읽을 수 없습니다.`n`n"
                 . "Outlook 오른쪽 위의 [새 Outlook] 스위치를 꺼서`n"
                 . "기존 Outlook 으로 바꾼 뒤 다시 눌러주세요.", "OUTLOOK Bar", 0x30)
            return r
        }
        if (state = "none") {
            MsgBox("Outlook 이 실행되고 있지 않습니다.`n`nOutlook 을 켠 뒤 다시 눌러주세요.", "OUTLOOK Bar", 0x30)
            return r
        }
        if (MsgBox("Outlook 에서 사서함 목록을 읽지 못했습니다.`n`n로그인한 본인 계정으로 진행할까요?", "OUTLOOK Bar", 0x4 | 0x30) != "Yes")
            return r
        r["ok"] := 1
        return r
    }

    pickBox := ""
    mnu := Menu()
    mnu.Add("사서함을 고르세요", (*) => "")
    mnu.Disable("사서함을 고르세요")
    mnu.Add()
    for lbl in labels
        mnu.Add(lbl, PickMailbox)
    mnu.Add()
    mnu.Add("취소", (*) => "")
    try mnu.Show()
    if (pickBox = "")
        return r
    r["ok"] := 1
    r["box"] := pickBox
    return r
}

; Review Daily — 최근 24시간 Inbox에서 답장 안 한 메일 리포트
RunReviewDaily(mailbox := "") {
    global busy
    if busy {
        Tip("이미 실행 중입니다 - 잠시만요...")
        return
    }
    if !FileExist(A_ScriptDir "\src\index.js") {
        MsgBox("src\index.js 를 찾을 수 없습니다.`nOUTLOOK_bar.ahk 는 Outlook 저장소 폴더 안에서 실행해야 합니다.", "OUTLOOK Bar", 0x30)
        return
    }
    box := mailbox
    if (box = "") {
        r := PickMailboxFlow()
        if !r["ok"]
            return
        box := r["box"]
    }
    EnvSet("OBAR_MBOXPICK", box)
    RunNodeShow(' /c node "src\index.js" --review-daily',
        "Review Daily 생성 중" (box != "" ? " (" box ")" : "") "... (10~30초, 완료되면 브라우저가 열립니다)")
}

RunFlaggedSummary(mailbox := "") {
    global busy, pickBox, pickMap
    if busy {
        Tip("이미 실행 중입니다 - 잠시만요...")
        return
    }
    if !FileExist(A_ScriptDir "\src\index.js") {
        MsgBox("src\index.js 를 찾을 수 없습니다.`nOUTLOOK_bar.ahk 는 Outlook 저장소 폴더 안에서 실행해야 합니다.", "OUTLOOK Bar", 0x30)
        return
    }

    box := mailbox
    if (box = "") {
        r := PickMailboxFlow()
        if !r["ok"]
            return
        box := r["box"]
    }

    ; ---- 폴더 선택: Inbox 하위 폴더가 있으면 전체/일부를 골라 요약한다 ----
    ; 한글 사서함·폴더 이름이 명령줄에서 깨지지 않도록 사서함은 환경변수로 전달
    EnvSet("OBAR_MBOXPICK", box)
    Tip("폴더 목록 확인 중...", 0)
    fout := CaptureNode(' /c node "src\index.js" --list-folders')
    HideTip()
    fspecs := []
    fdisps := []
    pos := 1
    while (hit := RegExMatch(fout, "m)^FLD=([^\t\r\n]+)\t([^\r\n]*)$", &fm, pos)) {
        pos := hit + StrLen(fm[0]) + 1
        fspecs.Push(fm[1])
        fdisps.Push(Trim(fm[2], " `t"))
    }
    if (fspecs.Length = 0) {
        ; 정상이라면 최소한 Inbox 한 줄은 온다 — 아무것도 없으면 폴더 목록 실패.
        ; 조용히 Inbox만 요약하지 말고 원인을 알려준다 (하위 폴더를 건너뛰는 문제의 원인 표시)
        ferr := ""
        if RegExMatch(fout, "m)^ERR=([^\r\n]*)", &fe)
            ferr := Trim(fe[1], " `t")
        if (ferr != "")
            fmsg := "폴더 목록을 읽지 못했습니다.`n(" ferr ")`n`n하위 폴더 없이 Inbox만 요약할까요?"
        else
            fmsg := "폴더 목록을 확인하지 못했습니다.`n프로그램 파일이 옛 버전이거나 설치가 불완전할 수 있습니다.`n`n[Update OUTLOOK Bar.bat] 를 한 번 실행하면 최신으로 맞춰집니다.`n`n일단 Inbox만 요약할까요?"
        if (MsgBox(fmsg, "OUTLOOK Bar", 0x4 | 0x30) != "Yes")
            return
        StartSummary(box, [])
        return
    }
    if (fspecs.Length = 1) {
        StartSummary(box, [])   ; 하위 폴더 없음 - Inbox만 (정상)
        return
    }
    ShowFolderPick(box, fspecs, fdisps)
}

; 폴더 선택 창 - 체크한 폴더만 요약 (기본 = 전체 체크)
ShowFolderPick(box, fspecs, fdisps) {
    global fg
    try {
        if IsObject(fg)
            fg.Destroy()
    }
    fg := Gui("+AlwaysOnTop +ToolWindow", "요약할 폴더 선택" (box != "" ? " - " box : ""))
    fg.SetFont("s10", "Segoe UI")
    fg.Add("Text", "xm", "요약할 폴더를 체크하세요 (기본 = 전체):")
    cbs := []
    for i, d in fdisps {
        cb := fg.Add("CheckBox", "xm Checked", d)
        cbs.Push(cb)
    }
    bAll := fg.Add("Button", "xm w90", "전체 선택")
    bAll.OnEvent("Click", (*) => FpSetAll(cbs, 1))
    bNone := fg.Add("Button", "x+8 w90", "전체 해제")
    bNone.OnEvent("Click", (*) => FpSetAll(cbs, 0))
    bGo := fg.Add("Button", "xm w110 Default", "요약 시작")
    bGo.OnEvent("Click", FpGo.Bind(box, fspecs, cbs))
    bCx := fg.Add("Button", "x+8 w90", "취소")
    bCx.OnEvent("Click", (*) => fg.Destroy())
    fg.OnEvent("Close", (*) => fg.Destroy())
    fg.Show()
}

FpSetAll(cbs, v) {
    for cb in cbs
        cb.Value := v
}

FpGo(box, fspecs, cbs, *) {
    global fg
    sel := []
    for i, cb in cbs {
        if cb.Value
            sel.Push(fspecs[i])
    }
    if (sel.Length = 0) {
        Tip("폴더를 하나 이상 체크해 주세요")
        return
    }
    fg.Destroy()
    StartSummary(box, sel)
}

; 실제 리포트 생성 실행 (sel: 폴더 spec 배열, 비면 Inbox만)
StartSummary(box, sel) {
    EnvSet("OBAR_MBOXPICK", box)
    cmd := ' /c node "src\index.js" --flagged-summary'
    if (sel.Length > 0) {
        tmp := A_Temp "\outlookbar_folders.txt"
        try FileDelete(tmp)
        try {
            f := FileOpen(tmp, "w", "UTF-8")
            for s in sel
                f.WriteLine(s)
            f.Close()
            cmd .= ' --folders-file "' tmp '"'
        }
    }
    RunNodeShow(cmd, "Flagged Summary 생성 중" (box != "" ? " (" box ")" : "") (sel.Length > 1 ? " - 폴더 " sel.Length "개" : "") "... (10~60초, 완료되면 브라우저가 열립니다)")
}

; node 실행 공통 — 로딩 표시를 유지하고, 실패하면 출력 마지막 줄(실제 원인)을 그대로 보여준다
RunNodeShow(cmd, waitMsg) {
    global busy
    busy := true
    Tip(waitMsg, 0)
    otmp := A_Temp "\outlookbar_run.txt"
    try FileDelete(otmp)
    ec := RunWait(A_ComSpec cmd ' > "' otmp '" 2>&1', A_ScriptDir, "Hide")
    HideTip()
    busy := false
    if (ec = 0)
        return
    if (ec = 9009) {
        MsgBox("Node.js가 설치되어 있지 않습니다.`n`n[Update OUTLOOK Bar.bat] 를 한 번 실행하면 자동으로 설치됩니다.", "OUTLOOK Bar", 0x30)
        return
    }
    detail := ""
    try {
        raw := FileRead(otmp, "UTF-8")
        keep := []
        for l in StrSplit(raw, "`n", "`r") {
            t := Trim(l, " `t")
            if (t != "")
                keep.Push(t)
        }
        i := keep.Length - 5
        if (i < 1)
            i := 1
        while (i <= keep.Length) {
            detail .= keep[i] "`n"
            i += 1
        }
        if (StrLen(detail) > 700)
            detail := SubStr(detail, -700)
    }
    emsg := "리포트 생성에 실패했습니다 (코드 " ec ")."
    if (detail != "")
        emsg .= "`n`n[원인]`n" detail
    emsg .= "`n- Outlook 이 켜져 있어야 합니다 (메일은 Outlook 에서 직접 읽습니다).`n- 방금 Outlook 을 켰다면 잠시 뒤 다시 눌러주세요."
    MsgBox(emsg, "OUTLOOK Bar", 0x30)
}

PickMailbox(name, pos, mnu) {
    global pickBox, pickMap
    pickBox := pickMap.Has(name) ? pickMap[name] : ""
}

; 노드 명령 출력을 임시파일로 받아 문자열로 돌려준다
CaptureNode(cmd) {
    tmp := A_Temp "\outlookbar_out.txt"
    try FileDelete(tmp)
    try RunWait(A_ComSpec cmd ' > "' tmp '"', A_ScriptDir, "Hide")
    out := ""
    try out := FileRead(tmp, "UTF-8")
    try FileDelete(tmp)
    return out
}


; ===== UPDATE 버튼 =====
; 자동 업데이트(켤 때 + 2시간마다)를 기다리지 않고 지금 바로 실행한다.
UpdateNow(*) {
    global AUTO_UPDATE, busy
    if busy {
        Tip("이미 실행 중입니다 - 잠시만요...")
        return
    }
    if !AUTO_UPDATE {
        MsgBox("여기는 개발용 원본 폴더라 업데이트를 받지 않습니다.`n(직원 PC에서는 정상 작동합니다)", "OUTLOOK Bar", 0x40)
        return
    }
    Tip("업데이트 확인 중...", 0)
    CheckUpdate(false, true)   ; 새 버전이면 교체 후 재시작되므로 아래로 안 온다
    HideTip()
}

; ===== 프로그램 파일(src) GitHub 자동 업데이트 =====
; why: 예전에는 회사 폴더(SharePoint) 복사에만 의존해서, 그게 안 되면 조용히 옛 버전으로 남았다.
; 바가 스스로 최신이 되는 것과 똑같은 방식으로 프로그램 파일도 받아온다.
; 받은 내용이 비었거나 GitHub 오류 페이지면 건너뛰고 기존 파일을 그대로 둔다 (망가뜨리지 않음).
RefreshSrcFromGitHub() {
    global SRC_BASE, SRC_FILES
    if InStr(A_ScriptDir, "GitHub")          ; 관리자 원본 폴더는 절대 덮어쓰지 않는다
        return 0
    dir := A_ScriptDir "\src"
    if !DirExist(dir) {
        try DirCreate(dir)
    }
    n := 0
    Loop SRC_FILES.Length {
        fn := SRC_FILES[A_Index]
        body := HttpGet(SRC_BASE fn "?v=" A_TickCount)
        if (body = "" || StrLen(body) < 50)
            continue
        if (SubStr(body, 1, 4) = "404:" || InStr(body, "<!DOCTYPE html"))
            continue
        path := dir "\" fn
        cur := ""
        try cur := FileRead(path, "UTF-8")
        if (NormTxt(cur) == NormTxt(body))
            continue
        try {
            f := FileOpen(path, "w", "UTF-8-RAW")
            f.Write(body)
            f.Close()
            n += 1
        }
    }
    return n
}

; ===== 프로그램 파일(src) 버전 대조 =====
; 바는 GitHub 에서 자동으로 최신이 되지만 프로그램 파일은 회사 폴더에서 복사돼 온다.
; 둘이 어긋나면 예전에는 아무 표시 없이 기능만 조용히 빠졌다 — 이제 바 제목에 계속 표시한다.
CheckSrcVersion(popup := false) {
    global srcVer, srcWarned, APPVER, DATEVER, g
    ; why: 원인이 둘이라 안내가 달라야 한다 — 파일이 옛 버전인가, node 가 안 도는가
    srcVer := ""
    cause := ""
    if !FileExist(A_ScriptDir "\src\index.js") {
        cause := "noprog"                       ; 프로그램 자체가 없음
    } else if !FileExist(A_ScriptDir "\src\srcver.js") {
        cause := "old"                          ; srcver.js 는 v23 부터 있음 → 그 이전 버전
    } else {
        out := CaptureNode(' /c node "src\srcver.js"')
        if RegExMatch(out, "SRCVER=(\d+)", &m)
            srcVer := m[1]
        if (srcVer = "")
            cause := "node"                     ; 파일은 있는데 실행이 안 됨 = node 문제
        else if (srcVer + 0 != APPVER)
            cause := "old"
    }

    title := "Outlook (" DATEVER ") v" APPVER
    if (cause = "old")
        title .= "  [프로그램 " (srcVer != "" ? "v" srcVer " " : "") "옛버전!]"
    else if (cause = "node")
        title .= "  [Node 문제!]"
    else if (cause = "noprog")
        title .= "  [프로그램 없음!]"
    try WinSetTitle(title, "ahk_id " g.Hwnd)

    if (cause = "" || !popup || srcWarned)
        return
    srcWarned := true
    if (cause = "node") {
        MsgBox("프로그램은 있는데 실행이 되지 않습니다 (Node 문제).`n`n"
             . "[Update OUTLOOK Bar.bat] 를 한 번 실행해 주세요 — Node 를 자동으로 설치합니다.", "OUTLOOK Bar", 0x30)
    } else if (cause = "noprog") {
        MsgBox("프로그램 파일이 없습니다.`n`n"
             . "[Update OUTLOOK Bar.bat] 를 한 번 실행해 주세요.", "OUTLOOK Bar", 0x30)
    } else {
        MsgBox("프로그램 파일이 옛 버전입니다.`n`n"
             . "바 = v" APPVER " / 프로그램 = " (srcVer != "" ? "v" srcVer : "v22 이전") "`n`n"
             . "이 상태에서는 새 기능(폴더 선택 등)이 동작하지 않습니다.`n`n"
             . "[Update OUTLOOK Bar.bat] 를 한 번 실행해 주세요.`n"
             . "실행해도 그대로면 회사 폴더가 아직 갱신되지 않은 것이니 Brian 에게 알려주세요.", "OUTLOOK Bar", 0x30)
    }
}

; ===== 자동 업데이트 =====
; ① 회사 공유 폴더에서 프로그램 파일(src\*.js 등) 최신화  ② GitHub 에서 바 파일 최신화 → 바뀌었으면 재시작
CheckUpdate(silent := true, force := false) {
    global UPDATE_URL, ini, APPVER
    ; 프로그램 파일(src) 동기화는 쿨다운과 상관없이 항상 먼저 한다.
    ; (예전: 바가 새 버전으로 바뀌며 재시작하면 2분 쿨다운에 걸려 src 를 못 받고
    ;  2시간 뒤에나 받았다 → "새 바 + 옛 프로그램" 상태로 폴더 목록이 실패했다)
    RefreshProgram()               ; 회사 폴더: config.json·서명 등
    srcN := RefreshSrcFromGitHub() ; GitHub: 프로그램 파일 (이쪽이 최신 기준)
    CheckSrcVersion(true)          ; 대조 — 팝업은 srcWarned 로 1회만

    ; UPDATE 버튼으로 직접 누른 경우(force) 2분 쿨다운을 무시한다
    last := IniRead(ini, "update", "lastreload", "")
    if (!force && last != "" && DateDiff(A_Now, last, "Seconds") < 120)
        return

    remote := HttpGet(UPDATE_URL "?v=" A_TickCount)
    if (remote = "" || StrLen(remote) < 800 || !InStr(remote, "OUTLOOK Desktop Bar") || !InStr(remote, "#Requires AutoHotkey")) {
        if !silent {
            HideTip()
            MsgBox("업데이트 확인에 실패했습니다.`n인터넷 연결을 확인하고 다시 눌러주세요.", "OUTLOOK Bar", 0x30)
        }
        return
    }
    cur := ""
    try cur := FileRead(A_ScriptFullPath, "UTF-8")
    if (NormTxt(remote) == NormTxt(cur)) {
        if !silent {
            HideTip()
            if (srcN > 0)
                MsgBox("프로그램 파일 " srcN "개를 최신으로 받았습니다.`n`n이제 최신 상태입니다 (v" APPVER ").", "OUTLOOK Bar", 0x40)
            else
                MsgBox("이미 최신입니다 (v" APPVER ").", "OUTLOOK Bar", 0x40)
        }
        return
    }
    try FileCopy(A_ScriptFullPath, A_Temp "\OUTLOOK_bar.bak.ahk", true)
    try {
        f := FileOpen(A_ScriptFullPath, "w", "UTF-8-RAW")
        f.Write(remote)
        f.Close()
    } catch {
        if !silent {
            HideTip()
            MsgBox("업데이트 저장에 실패했습니다.`n바를 껐다가 다시 켠 뒤 눌러주세요.", "OUTLOOK Bar", 0x30)
        }
        return
    }
    IniWrite(A_Now, ini, "update", "lastreload")
    Tip("OUTLOOK 바 업데이트됨 - 잠시만요...")
    Sleep(800)
    Reload()
}

; 회사 공유 폴더(Admin\Automation\Outlook)에서 프로그램 파일을 조용히 최신화
RefreshProgram() {
    if InStr(A_ScriptDir, "GitHub")
        return
    ps := 'powershell -NoProfile -ExecutionPolicy Bypass -Command "'
        . "$mids=@('Timeless 042026 - Documents\Admin\Automation\Outlook','Admin\Automation\Outlook','Automation\Outlook');"
        . "$roots=@($env:OneDriveCommercial,$env:OneDrive);"
        . "Get-ChildItem $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | ForEach-Object { $roots += $_.FullName };"
        . "$src='';"
        . "foreach($r in $roots){ if(-not $r){continue}; foreach($m in $mids){ $p=Join-Path $r $m; if(Test-Path (Join-Path $p 'src\index.js')){ $src=$p; break } }; if($src){break} };"
        . "if($src){ robocopy $src '" A_ScriptDir "' /E /XD node_modules reports .git .claude docs /XF tokens.json state.json log.txt 'flagged-cache*.json' OUTLOOK_bar.ini OUTLOOK_bar.ahk | Out-Null }"
        . '"'
    try RunWait(A_ComSpec ' /c ' ps, , "Hide")
}

HttpGet(url) {
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", url, false)
        whr.SetTimeouts(3000, 3000, 3000, 5000)
        whr.SetRequestHeader("Cache-Control", "no-cache")
        whr.SetRequestHeader("Pragma", "no-cache")
        whr.Send()
        if (whr.Status = 200)
            return whr.ResponseText
    }
    return ""
}

NormTxt(s) {
    s := StrReplace(s, Chr(0xFEFF), "")
    s := StrReplace(s, "`r", "")
    return Trim(s, " `t`n")
}

; ===== 로딩/알림 표시 =====
; 마우스 위치에 고정되던 ToolTip 대신, 바 바로 밑에 붙는 노란 상태창을 쓴다.
; 바를 옮기면 같이 따라오고(150ms), dur=0 이면 HideTip() 할 때까지 계속 보인다.
Tip(msg, dur := 2500) {
    global stickyMsg, tipSeq
    tipSeq += 1
    if (dur <= 0) {
        stickyMsg := msg
        ShowTipWin(msg)
        return
    }
    ShowTipWin(msg)
    SetTimer(TipExpire.Bind(tipSeq), -dur)
}

TipExpire(seq) {
    global tipSeq, stickyMsg
    if (seq != tipSeq)
        return
    if (stickyMsg != "")
        ShowTipWin(stickyMsg)   ; 잠깐 알림이 끝나면 진행 중 메시지로 복귀
    else
        HideTip()
}

ShowTipWin(msg) {
    global tipG, scale, g
    try {
        if IsObject(tipG)
            tipG.Destroy()
    }
    fs := Round(11 * scale)
    if (fs < 10)
        fs := 10
    tipG := Gui("+AlwaysOnTop -Caption +ToolWindow +Border", "OutlookBarTip")
    tipG.BackColor := "FFD54A"
    tipG.MarginX := 12
    tipG.MarginY := 8
    tipG.SetFont("s" fs " Bold", "Segoe UI")
    tipG.Add("Text", "c202020", msg)
    bx := 0
    by := 0
    bw := 0
    bh := 0
    try WinGetPos(&bx, &by, &bw, &bh, "ahk_id " g.Hwnd)
    tipG.Show("x" bx " y" (by + bh + 4) " NoActivate")
    PlaceTip()
    SetTimer(PlaceTip, 150)
}

; 상태창을 바 바로 밑 가운데에 붙인다 (바가 움직이면 타이머가 계속 따라붙게)
PlaceTip() {
    global g, tipG
    if !IsObject(tipG) {
        SetTimer(PlaceTip, 0)
        return
    }
    try {
        WinGetPos(&bx, &by, &bw, &bh, "ahk_id " g.Hwnd)
        WinGetPos(, , &tw, &th, "ahk_id " tipG.Hwnd)
        nx := bx + (bw - tw) // 2
        if (nx < 0)
            nx := 0
        ny := by + bh + 4
        if (ny + th > A_ScreenHeight - 4)
            ny := by - th - 4
        WinMove(nx, ny, , , "ahk_id " tipG.Hwnd)
    } catch {
        SetTimer(PlaceTip, 0)
    }
}

HideTip() {
    global tipG, stickyMsg, tipSeq
    tipSeq += 1
    stickyMsg := ""
    SetTimer(PlaceTip, 0)
    try {
        if IsObject(tipG)
            tipG.Destroy()
    }
    tipG := ""
}
