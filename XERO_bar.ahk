#Requires AutoHotkey v2.0
#SingleInstance Force

; ================= XERO Desktop Bar =================
; 버튼 = 크롬으로 단축키(F13~F20) 전송 → Tampermonkey가 받아 현재 Xero 탭에서 실행.
; EDIT: 보여줄 버튼만 체크 → SAVE(저장) / X(취소). 선택은 저장돼 다음에도 유지.

VER := "06/07/26"                       ; 제목 표시 날짜
ini := A_ScriptDir "\XERO_bar.ini"      ; 버튼 선택 저장

tools := [
    {label:"Xero Reset",     c:"8B0000", key:"F13", id:"xeroreset"},
    {label:"Increase Apply", c:"1F4E78", key:"F14", id:"increaseapply"},
    {label:"Stripe Off",     c:"6C63FF", key:"F15", id:"stripeoff"},
    {label:"App4Sending",    c:"2E7D32", key:"F16", id:"app4sending"},
    {label:"Stripe+Sending", c:"00695C", key:"F17", id:"stripesending"},
    {label:"Xero 20",        c:"C8511B", key:"F18", id:"xero20"},
    {label:"Xero ALL 20",    c:"C8511B", key:"F19", id:"xeroall20"},
    {label:"XERO help",      c:"555555", key:"F20", id:"xerohelp"}
]

enabled := Map()
for t in tools
    enabled[t.id] := (IniRead(ini, "tools", t.id, "1") = "1")

editMode := false
posX := 20
posY := 150
g := ""
checks := Map()

; 첫 실행 시 바탕화면에 "XERO Bar" 아이콘 자동 생성 (다음부턴 그걸로 켜기)
; auto-create a "XERO Bar" desktop icon on first run (use it to re-open later)
try {
    _lnk := A_Desktop "\XERO Bar.lnk"
    if !FileExist(_lnk)
        FileCreateShortcut(A_ScriptFullPath, _lnk, A_ScriptDir)
}

Build()

Build() {
    global g, tools, VER, enabled, editMode, posX, posY, checks
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
    g := Gui("+AlwaysOnTop +ToolWindow -MinimizeBox", "XERO (" VER ")")
    g.BackColor := "1F2A38"
    g.SetFont("s10 Bold", "Segoe UI")
    g.OnEvent("Close", (*) => ExitApp())

    if editMode {
        sv := g.Add("Text", "w150 h24 Center 0x200 Background1E88E5 cWhite", "SAVE")
        sv.OnEvent("Click", (*) => SaveEdit())
        cx := g.Add("Text", "x+4 yp w26 h24 Center 0x200 BackgroundB71C1C cWhite", "X")
        cx.OnEvent("Click", (*) => CancelEdit())
        for t in tools {
            cb := g.Add("CheckBox", "xm y+8 w180 " (enabled[t.id] ? "Checked" : "") " cWhite", t.label)
            checks[t.id] := cb
        }
    } else {
        et := g.Add("Text", "w180 h24 Center 0x200 Background455A64 cWhite", "EDIT")
        et.OnEvent("Click", (*) => EnterEdit())
        for t in tools {
            if !enabled[t.id]
                continue
            b := g.Add("Text", "w180 h32 Center 0x200 Background" t.c " cWhite", t.label)
            b.OnEvent("Click", SendKey.Bind(t.key))
        }
    }
    g.Show("x" posX " y" posY " AutoSize")
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
