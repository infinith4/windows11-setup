#Requires AutoHotkey v2.0

#HotIf WinActive("ahk_class CabinetWClass")
^+m::
{
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    
    ; パスを取得
    Send("!d")
    Sleep(300)
    Send("^c")
    ClipWait(1)
    path := A_Clipboard
    Send("{Escape}")
    Sleep(500)
    
    ; 新しいタブを開く
    Send("^t")
    Sleep(900)  ; 待ち時間を長く
    
    ; パスを入力
    A_Clipboard := path
    Sleep(400)
    Send("!d")
    Sleep(500)  ; さらに長く
    Send("^a")
    Sleep(200)
    Send("^v")
    Sleep(400)
    Send("{Enter}")
    
    A_Clipboard := ClipSaved
}
#HotIf