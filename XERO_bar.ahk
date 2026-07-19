#Requires AutoHotkey v2.0
#SingleInstance Force

; ================= XERO Desktop Bar =================
; 버튼 = 크롬으로 단축키(F13~F22) 전송 → Tampermonkey가 받아 현재 Xero 탭에서 실행.
; EDIT: 보여줄 버튼만 체크 → SAVE(저장) / X(취소). 순서는 일반 화면에서 버튼을 위/아래로 드래그해 변경. 선택·순서는 저장돼 다음에도 유지.
; 크기 조절: 창 오른쪽 아래 코너를 마우스로 끌어서 늘리거나 줄이세요. 크기는 저장됩니다.

VER := "15/07/26"                       ; 기본 날짜(오프라인용). 켜지면 웹페이지와 같은 날짜를 읽어와 자동 표시.
PAGE_URL := "https://timeless15000.github.io/xero-apps/Xero_applications.html"  ; 제목 날짜 출처(웹페이지와 동일)
ini := A_ScriptDir "\XERO_bar.ini"      ; 버튼 선택 / 크기 저장

; ---- 자동 업데이트 ----
UPDATE_URL := "https://raw.githubusercontent.com/Timeless15000/xero-apps/main/XERO_bar.ahk"
; xero-apps(클론) 폴더에서 실행하면 '원본'이라 자동교체 끔 - 편집 중인 파일을 덮어쓰지 않도록.
AUTO_UPDATE := !InStr(A_ScriptFullPath, "xero-apps")

tools := [
    {label:"Xero Reset",     c:"8B0000", key:"F13", id:"xeroreset"},
    {label:"Increase Apply", c:"1F4E78", key:"F14", id:"increaseapply"},
    {label:"Stripe Off",     c:"6C63FF", key:"F15", id:"stripeoff"},
    {label:"Stripe Off+Save",c:"5E35B1", key:"F21", id:"stripeoffsave"},
    {label:"App4Sending",    c:"2E7D32", key:"F16", id:"app4sending"},
    {label:"Stripe+Sending", c:"00695C", key:"F17", id:"stripesending"},
    {label:"Xero 20",        c:"C8511B", key:"F18", id:"xero20"},
    {label:"Xero ALL 20",    c:"C8511B", key:"F19", id:"xeroall20"},
    {label:"XERO help",      c:"555555", key:"F20", id:"xerohelp"},
    {label:"Price Check",    c:"0097A7", key:"F22", id:"pricecheck"}
]

enabled := Map()
for t in tools
    enabled[t.id] := (IniRead(ini, "tools", t.id, "1") = "1")

LoadOrder()                             ; 저장된 순서(있으면)대로 tools 재정렬

editMode := false
posX := 20
posY := 150
scale := IniRead(ini, "cfg", "scale", "1") + 0    ; 크기 배율 (1 = 100%)
if (scale < 0.6)
    scale := 0.6
if (scale > 3)
    scale := 3
opacity := IniRead(ini, "cfg", "opacity", "100") + 0   ; 투명도 (100 = 불투명, 낮을수록 투명)
if (opacity < 20)
    opacity := 20
if (opacity > 100)
    opacity := 100
g := ""
checks := Map()
items := []                             ; 컨트롤 + 기본(100%) 좌표 저장
dragRows := []                          ; 일반 모드 버튼 목록(드래그로 순서 변경)

; 첫 실행 시 바탕화면에 "XERO Bar" 아이콘 자동 생성 (다음부턴 그걸로 켜기)
try {
    _lnk := A_Desktop "\XERO Bar.lnk"
    if !FileExist(_lnk)
        FileCreateShortcut(A_ScriptFullPath, _lnk, A_ScriptDir)
}

; 다른 위치에서 돌던 예전 XERO 바가 있으면 닫기 (바 중복 방지)
SetTitleMatchMode(2)
try {
    for _hw in WinGetList("XERO (")
        try WinClose("ahk_id " _hw)
}

Build()

if AUTO_UPDATE {
    SetTimer(() => CheckUpdate(true), -3000)              ; 켠 뒤 3초 후 1회 확인
    SetTimer(() => CheckUpdate(true), 2 * 60 * 60 * 1000) ; 이후 2시간마다 확인
}

; 제목 날짜를 웹페이지와 똑같이 유지 (켠 직후 1회 + 2시간마다)
SetTimer(() => RefreshVer(), -1500)
SetTimer(() => RefreshVer(), 2 * 60 * 60 * 1000)

Build() {
    global g, tools, VER, enabled, editMode, posX, posY, checks, scale, items, opacity, dragRows
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
    dragRows := []
    g := Gui("+AlwaysOnTop +Resize +ToolWindow -MaximizeBox -MinimizeBox", "XERO (" VER ")")
    g.BackColor := "1F2A38"
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
            b.OnEvent("Click", SendKey.Bind(t.key))   ; 드래그 훅 미작동 시 대비(클릭=실행)
            AddItem(b, M, y, CW, 32)
            dragRows.Push({hwnd: b.Hwnd, ctrl: b, id: t.id, key: t.key, label: t.label, baseY: y})
            y += 32 + gap
        }
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

; 코너를 끌면 폭에 맞춰 바 전체를 확대/축소
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

; ================= 드래그로 순서 변경 =================
; 일반 화면에서 버튼을 위/아래로 끌어 순서 변경. 살짝(반 칸 미만) 누르면 그대로 실행.
; 마우스가 버튼 위에 있을 때만 LButton 훅이 켜져 → 슬라이더·SAVE·EDIT 등은 정상 동작.
IsOverToolButton() {
    global g, dragRows
    if !IsObject(g)
        return false
    if !dragRows.Length
        return false
    MouseGetPos(, , &win, &ch, 2)
    if (win != g.Hwnd)
        return false
    for r in dragRows
        if (r.hwnd = ch)
            return true
    return false
}

DragReorder() {
    global g, tools, dragRows, scale
    MouseGetPos(&sx, &sy, &win, &ch, 2)
    if (win != g.Hwnd)
        return
    di := 0
    for i, r in dragRows
        if (r.hwnd = ch) {
            di := i
            break
        }
    if (di = 0)
        return
    n := dragRows.Length
    pitch := (n >= 2) ? Abs(dragRows[2].baseY - dragRows[1].baseY) * scale : 38 * scale
    if (pitch < 1)
        pitch := 38 * scale
    ; 드래그: 슬롯 단위로 자리 바꾸기 (겹침/잔상 없이 깔끔하게)
    order := []                            ; 현재 표시 순서 (dragRows 엔트리 복사)
    for r in dragRows
        order.Push(r)
    dragRows[di].ctrl.GetPos(&bx, &by0, &bw0, &bh0)   ; 버튼 x(열) - 세로만 바꿈
    dragKey := dragRows[di].key
    curIdx := di
    moved := false
    while GetKeyState("LButton", "P") {
        MouseGetPos(, &my)
        target := di + Round((my - sy) / pitch)
        if (target < 1)
            target := 1
        if (target > n)
            target := n
        if (target != curIdx) {            ; 슬롯이 바뀔 때만 전체 재배치 (깔끔)
            moved := true
            entry := order.RemoveAt(curIdx)
            order.InsertAt(target, entry)
            Loop n
                order[A_Index].ctrl.Move(bx, Round(dragRows[A_Index].baseY * scale))
            curIdx := target
        }
        Sleep(10)
    }
    if (!moved) {                          ; 안 움직였으면 클릭(실행)
        SendKey(dragKey)
        return
    }
    visIds := []                           ; 최종 표시 순서
    for e in order
        visIds.Push(e.id)
    visSet := Map()                        ; 전체 tools 에 반영(숨긴 버튼 자리는 유지)
    byId := Map()
    for r in dragRows
        visSet[r.id] := true
    for t in tools
        byId[t.id] := t
    vi := 1
    ordered := []
    for t in tools {
        if visSet.Has(t.id) {
            ordered.Push(byId[visIds[vi]])
            vi += 1
        } else {
            ordered.Push(t)
        }
    }
    tools := ordered
    SaveOrder()
    Build()
}

; 저장된 순서(있으면)대로 tools 재정렬. 저장 순서에 없는(새로 추가된) 버튼은 뒤에 붙임.
LoadOrder() {
    global tools, ini
    ord := IniRead(ini, "cfg", "order", "")
    if (ord = "")
        return
    byId := Map()
    for t in tools
        byId[t.id] := t
    ordered := []
    seen := Map()
    for id in StrSplit(ord, ",") {
        id := Trim(id)
        if (id != "" && byId.Has(id) && !seen.Has(id)) {
            ordered.Push(byId[id])
            seen[id] := true
        }
    }
    for t in tools
        if !seen.Has(t.id) {
            ordered.Push(t)
            seen[t.id] := true
        }
    tools := ordered
}

; 현재 tools 순서를 .ini 에 저장 (id 를 콤마로 연결)
SaveOrder() {
    global tools, ini
    ids := ""
    for t in tools
        ids .= (ids = "" ? "" : ",") t.id
    IniWrite(ids, ini, "cfg", "order")
}

SendKey(key, *) {
    hwnd := WinExist("ahk_exe chrome.exe")
    if hwnd {
        WinActivate(hwnd)
        if WinWaitActive(hwnd, , 1)
            Send("{" key "}")
    } else {
        MsgBox("Open a Xero page in Chrome first.", "XERO", 0x40)
    }
}

; ================= 자동 업데이트 =================
; 켤 때 + 2시간마다 GitHub에서 최신 바를 확인.
; 새 버전이면 이 파일을 스스로 교체하고 리로드. .ini(설정)는 건드리지 않음.
CheckUpdate(silent := true) {
    global UPDATE_URL, ini
    ; 리로드 직후 반복되는 것 방지: 직전 리로드 후 2분 이내면 건너뜀
    last := IniRead(ini, "update", "lastreload", "")
    if (last != "" && DateDiff(A_Now, last, "Seconds") < 120)
        return
    remote := HttpGet(UPDATE_URL "?v=" A_TickCount)
    ; 다운로드 검증 (네트워크 오류/404/반쪽짜리면 교체 안 함)
    if (remote = "" || StrLen(remote) < 800 || !InStr(remote, "XERO Desktop Bar") || !InStr(remote, "#Requires AutoHotkey")) {
        if !silent
            Tip("업데이트 확인 실패 - 인터넷/GitHub 확인")
        return
    }
    cur := ""
    try cur := FileRead(A_ScriptFullPath, "UTF-8")
    if (NormTxt(remote) == NormTxt(cur)) {
        if !silent
            Tip("이미 최신이에요")
        return
    }
    ; 새 버전 → 백업 후 교체 → 리로드
    try FileCopy(A_ScriptFullPath, A_ScriptDir "\XERO_bar.bak.ahk", true)
    try {
        f := FileOpen(A_ScriptFullPath, "w", "UTF-8-RAW")
        f.Write(remote)
        f.Close()
    } catch {
        if !silent
            Tip("업데이트 저장 실패")
        return
    }
    IniWrite(A_Now, ini, "update", "lastreload")
    Tip("XERO 바 업데이트됨 - 잠시만요...")
    Sleep(800)
    Reload()
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
    s := StrReplace(s, Chr(0xFEFF), "")   ; BOM 제거
    s := StrReplace(s, "`r", "")          ; CRLF → LF 통일
    return Trim(s, " `t`n")
}

Tip(msg) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -2500)
}

; ================= 제목 날짜 = 웹페이지와 동일 =================
; 웹페이지 상단 날짜 = GitHub Pages 배포시각(document.lastModified).
; 바도 그 값(HTTP Last-Modified)을 읽어 로컬시간으로 바꿔 제목에 표시 → 항상 같은 날짜.
RefreshVer() {
    global VER, g, PAGE_URL
    d := GetPageDate(PAGE_URL)
    if (d != "" && d != VER) {
        VER := d
        if IsObject(g)
            try g.Title := "XERO (" VER ")"
    }
}

GetPageDate(url) {
    lm := ""
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", url "?v=" A_TickCount, false)
        whr.SetTimeouts(3000, 3000, 3000, 5000)
        whr.SetRequestHeader("Cache-Control", "no-cache")
        whr.SetRequestHeader("Pragma", "no-cache")
        whr.Send()
        if (whr.Status = 200)
            lm := whr.GetResponseHeader("Last-Modified")
    }
    if (lm = "")
        return ""
    ; 예: "Wed, 15 Jul 2026 08:28:33 GMT"
    if !RegExMatch(lm, "i)(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})\s+(\d{2}):(\d{2}):(\d{2})", &mm)
        return ""
    mon := Map("Jan","01","Feb","02","Mar","03","Apr","04","May","05","Jun","06","Jul","07","Aug","08","Sep","09","Oct","10","Nov","11","Dec","12")
    if !mon.Has(mm[2])
        return ""
    gmt := mm[3] . mon[mm[2]] . Format("{:02}", mm[1]+0) . mm[4] . mm[5] . mm[6]   ; YYYYMMDDHHMISS (GMT)
    off := DateDiff(A_Now, A_NowUTC, "Seconds")                                    ; 로컬-UTC 오프셋(초)
    loc := DateAdd(gmt, off, "Seconds")                                            ; 로컬 시간으로 변환
    return SubStr(loc, 7, 2) . "/" . SubStr(loc, 5, 2) . "/" . SubStr(loc, 3, 2)   ; DD/MM/YY
}

; ---- 일반 화면에서 버튼을 위/아래로 드래그해 순서 변경 ----
; (마우스가 버튼 위에 있을 때만 켜짐 → 그 외 클릭은 평소대로 동작)
#HotIf IsOverToolButton()
LButton:: {
    DragReorder()
}
#HotIf
