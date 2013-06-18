#include <ScreenCapture.au3>
Local $answer = MsgBox(1, "把我放在PS2上面", "把鼠标放在输入框处，然后按一下空格，注意：测试时别乱动，不要按键盘，不要按鼠标，但鼠标可以移动，对了，测试完了，我会通知你的。你可以按一下Ctrl+Shift+q这个组合键来强行退出测试，Goodluck 亲！")

HotKeySet("^+q", "Terminate")

If $answer = 1 Then
   TestMain()
   MsgBox(0, "测试完成", "大功告成")
   Exit
EndIf


Func ReadAndInputFromFile($file)
   Local $iCount = 1
   Local $iFileE = 1
   ;循环读取文件每一行，并输入至输入框
   while 1
      
       Local $line = FileReadLine($file)
       If @error = -1 Then
          ExitLoop
       Endif

       Send("{ENTER}")
       Sleep(200)
       Send($line)
       Sleep(300)
      
       $iCount = $iCount + 1
       If $iCount >=30 Then
          $iCount = 1
          _ScreenCapture_Capture(@MyDocumentsDir & "\TestImage_" & $iFileE & ".jpg")
          $iFileE = $iFileE + 1
       EndIf
      
       Send("{ENTER}")
       Sleep(4500)
      
WEnd
EndFunc


Func TestMain()
   ;设置焦点为PS2输入框
     Sleep(1000)
     MouseClick("left")
    
   ;打开配置文件
   Local $file = FileOpen(@DesktopDir & "\1.txt", 0)
   If $file = -1 Then
       MsgBox(0, "出错了", "哥们儿，你是不是把  1.txt  这个玩意弄没了，记住，这个家伙要放在桌面上~")
       Exit
   EndIf

   ReadAndInputFromFile($file)
  
   ;关闭文件名柄
   FileClose($file)
EndFunc

Func Terminate()
    Exit 0
EndFunc