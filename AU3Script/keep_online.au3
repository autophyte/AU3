Local $answer = MsgBox(1, "����", "����԰�һ��Ctrl+Shift+q�����ϼ���ǿ���˳�����")

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