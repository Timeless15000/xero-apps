#Requires AutoHotkey v2.0
#SingleInstance Force

; ================= XERO Desktop Bar =================
; 버튼 = 크롬으로 단축키(F13~F20) 전송 → Tampermonkey가 받아 현재 Xero 탭에서 실행.
; EDIT: 보여줄 버튼만 체크 → SAVE(저장) / X(취소). 선택은 저장돼 다음에도 유지.
; 크기 조절: 창 오른쪽 아래 코너를 마우스로 끌어서 늘리거나 줄이세요. 크기는 저장됩니다.

VER := "06/07/26"                       ; 제목 표시 날짜
ini := A_ScriptDir "\XERO_bar.ini"      ; 버튼 선택 / 크기 저장

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

; 첫 실행 시 바탕화면에 "XERO Bar" 아이콘 자동 생성 (다음부턴 그걸로 켜기)
try {
    _lnk := A_Desktop "\XERO Bar.lnk"
    if !FileExist(_lnk)
        FileCreateShortcut(A_ScriptFullPath, _lnk, A_ScriptDir)
}

Build()

Build() {
    global g, tools, VER, enabled, editMode, posX, posY, checks, scale, items, opacity
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
            b.OnEvent("Click", SendKey.Bind(t.key))
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
