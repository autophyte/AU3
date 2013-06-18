#include <Color.au3>
#include <ScreenCapture.au3>
HotKeySet("^+q", "Terminate")


Local $cap_left_top
local $cap_right_bootom
Local $pos_left_top
local $pos_right_bootom
Local $pos_input
Local $strLine

Func SetTestRange()
   ;设定截图范围，当出现没有被列入屏蔽字的名字时，将会截图，用于后期判断
   MsgBox(0x40000, "设定截图位置", "将鼠标移到截图区域左上角，然后按空格键")
   $cap_left_top = MouseGetPos()
   MsgBox(0x40000, "设定截图位置", "将鼠标移到截图区域右下角，然后按空格键")
   $cap_right_bootom = MouseGetPos()

   ;设定完成按钮的区域，脚本用“完成按钮”颜色来确定名字是否可用
   ;设置搜索颜色范围，位置1为完成按钮的左上角，位置2为完成按钮的右下角，可以更广阔
   MsgBox(0x40000, "设定按钮位置", "将鼠标移到完成按钮左上角，然后按空格键")
   $pos_left_top = MouseGetPos()
   MsgBox(0x40000, "设定按钮位置", "将鼠标移到完成按钮右下角，然后按空格键")
   $pos_right_bootom = MouseGetPos()
 
   ;设定输入框的位置
   MsgBox(0x40000, "设定输入框位置", "请将鼠标移到输入框位置，按下空格键")
   $pos_input = MouseGetPos()
EndFunc


Func ReadAndInputFromFile($file, $file_r)
 
   Local $iLine = 1
   ;循环读取文件每一行，并输入至输入框
   while 1
       $strLine = FileReadLine($file)
       If @error = -1 Then
          ExitLoop
       Endif
       
        If StringLen($strLine) <= 1 Then
            $strLine = $strLine & " ";
        EndIf

       ;焦点设置在输入框内，并清空输入框内容
       MouseClick("left", $pos_input[0], $pos_input[1])
       Send(("^{a}"))
       Sleep(100)
       Send("{BS}")
     
       ;输入读取到的名称?
       Send($strLine)
        Send("{ENTER}")
       Sleep(2000)
     
       ;搜索颜色2780482，这个颜色是绿色，当完成按钮高亮，也就是名字可用时，范围内应该包含这个颜色
       ;可以找到这个颜色，此时表示名字没有被列入屏蔽字列表中，应该记录
       Local $coord = PixelSearch($pos_left_top[0], $pos_left_top[1], $pos_right_bootom[0], $pos_right_bootom[1], 2780482, 5)
      If Not @error Then
          ;截图
          _ScreenCapture_Capture(@MyDocumentsDir & "\CharacterNameImage_" & $iLine & ".jpg", $cap_left_top[0], $cap_left_top[1], $cap_right_bootom[0], $cap_right_bootom[1])

          ;记录到文件中
          FileWriteLine($file_r, $iLine & @TAB & $strLine & @TAB & "\CharacterNameImage_" & $iLine & ".jpg" & @CRLF)
          FileFlush($file_r)
       EndIf
     
       $iLine = $iLine + 1
   WEnd
EndFunc

Func TestMain()
   SetTestRange()
 
   ;打开配置文件
   Local $file = FileOpen(@DesktopDir & "\1.txt", 0)
   If $file = -1 Then
       MsgBox(0, "出错了", "哥们儿，你是不是把  1.txt  这个玩意弄没了，记住，这个家伙要放在桌面上~")
       Exit
   EndIf
 
   ;打开report文件
   Local $file_r = FileOpen(@MyDocumentsDir & "\CharacterNameReport.txt", 10)
   If $file = -1 Then
       MsgBox(0, "出错了", "哥们儿，具体原因我也不知道，反正就是我没有创建成功~")
       Exit
   EndIf
   FileWriteLine($file_r, "LineNO." & @TAB & "Name" & @TAB & "Picture" & @CRLF)

   ReadAndInputFromFile($file, $file_r)
 
   FileClose($file)
   FileClose($file_r)
EndFunc


Func Terminate()
   MsgBox(0, "测试中断", "当前测试进度：" & $strLine)
   Exit 0
EndFunc

TestMain()
MsgBox(0, "测试完成", "大功告成")