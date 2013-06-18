Local $answer = MsgBox(1, "测试", "你可以按一下Ctrl+Shift+q这个组合键来强行退出测试")

HotKeySet("^+q", "Terminate")

If $answer = 1 Then
   TestMain()
   Exit
EndIf


Func TestMain()
   while 1
       Sleep(5000)
       Send("{SPACE}")
       Sleep(5000)
       MouseClick("right")
     WEnd
EndFunc


Func Terminate()
    Exit 0
EndFunc