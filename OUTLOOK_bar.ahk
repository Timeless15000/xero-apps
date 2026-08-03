#Requires AutoHotkey v2.0
#SingleInstance Force

; ================= OUTLOOK Desktop Bar =================
; XERO 바와 같은 조작감의 데스크톱 바. 차이점: 버튼이 단축키를 보내는 대신
; 이 저장소의 Node 명령을 직접 실행한다 (Tampermonkey 불필요).
; - Flagged Summary: flag된 메일을 AI 요약과 함께 HTML 리포트로 정리해 브라우저로 연다.
;   지금 Outlook 화면에 보고 있는 사서함(공유·추가 사서함 포함)을 자동으로 대상으로 삼는다.
;   특정 사서함을 항상 보려면 tools 에 mbox 를 지정한 버튼을 두면 된다.
; EDIT: 보여줄 버튼만 체크 → SAVE / X(취소). 크기 조절: 창 오른쪽 아래 코너 드래그. 설정은 저장됨.

ini := A_ScriptDir "\OUTLOOK_bar.ini"

APPVER := 11                              ; 앱 버전 — 수정할 때마다 +1 (제목에 v6 처럼 표시)
DATEVER := "02/08/2026"                 ; 오프라인 기본값. 아래에서 파일 수정날짜로 자동 대체.
try DATEVER := FormatTime(FileGetTime(A_ScriptFullPath, "M"), "dd/MM/yyyy")  ; 이 파일 마지막 수정일 = 버전 날짜

; 공유 사서함 버튼을 늘리려면 {label, c, id:"영문id", mbox:"<메일주소>"} 한 줄을 추가하면 된다
tools := [
    {label:"Flagged Summary", c:"0F6CBD", id:"flaggedsummary", mbox:""}   ; 보고 있는 사서함 자동
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
checks := Map()
pickBox := ""
pickMap := Map()
items := []

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
            RunFlaggedSummary(t.HasOwnProp("mbox") ? t.mbox : "")
            return
        }
    }
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
        ; Outlook 에 붙어 있는 사서함 목록을 받아 직접 고르게 한다
        Tip("사서함 목록 확인 중...", 20000)
        out := CaptureNode(' /c node "src\index.js" --list-mailboxes')
        ToolTip()

        cur := ""
        if RegExMatch(out, "CUR=([^\r\n]*)", &mc)
            cur := Trim(mc[1], " `t")

        pickMap := Map()
        labels := []
        pos := 1
        while (hit := RegExMatch(out, "m)^MBOX=([^\t\r\n]+)\t?([^\r\n]*)$", &m, pos)) {
            pos := hit + StrLen(m[0]) + 1
            smtp := Trim(m[1], " `t")
            nm := Trim(m[2], " `t")
            lbl := (nm != "" && nm != smtp) ? nm "  (" smtp ")" : smtp
            if (smtp = cur)
                lbl := lbl "   ← 지금 보는 폴더"
            if !pickMap.Has(lbl) {
                pickMap[lbl] := smtp
                labels.Push(lbl)
            }
        }

        if (labels.Length = 0) {
            if (MsgBox("Outlook 에서 사서함 목록을 읽지 못했습니다.`n(Outlook 이 켜져 있는지 확인해 주세요)`n`n로그인한 본인 계정으로 요약할까요?", "OUTLOOK Bar", 0x4 | 0x30) != "Yes")
                return
        } else {
            pickBox := ""
            mnu := Menu()
            mnu.Add("요약할 사서함을 고르세요", (*) => "")
            mnu.Disable("요약할 사서함을 고르세요")
            mnu.Add()
            for lbl in labels
                mnu.Add(lbl, PickMailbox)
            mnu.Add()
            mnu.Add("취소", (*) => "")
            try mnu.Show()
            if (pickBox = "")
                return
            box := pickBox
        }
    }

    busy := true
    Tip("Flagged Summary 생성 중" (box ? " (" box ")" : "") "... (10~60초, 완료되면 브라우저가 열립니다)", 60000)
    cmd := ' /c node "src\index.js" --flagged-summary'
    if (box != "")
        cmd .= ' --mailbox "' box '"'
    ec := RunWait(A_ComSpec cmd, A_ScriptDir, "Hide")
    ToolTip()
    busy := false
    if (ec = 9009) {
        MsgBox("Node.js가 설치되어 있지 않습니다.`n`nhttps://nodejs.org 에서 LTS 버전을 설치한 뒤 다시 눌러주세요.", "OUTLOOK Bar", 0x30)
    } else if (ec != 0) {
        MsgBox("리포트 생성에 실패했습니다 (코드 " ec ").`n`n- 로그인이 만료됐거나 권한이 새로 필요하면:`n  Outlook 폴더의 ""Login again.bat"" 을 실행하세요.`n- 다른 사람/다른 회사 계정의 사서함은 접근 권한이 있어야 합니다.`n- 자세한 내용은 log.txt 를 확인하세요.", "OUTLOOK Bar", 0x30)
    }
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


Tip(msg, dur := 2500) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -dur)
}
