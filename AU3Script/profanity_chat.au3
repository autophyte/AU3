#include <ScreenCapture.au3>
Local $answer = MsgBox(1, "���ҷ���PS2����", "������������򴦣�Ȼ��һ�¿ո�ע�⣺����ʱ���Ҷ�����Ҫ�����̣���Ҫ����꣬���������ƶ������ˣ��������ˣ��һ�֪ͨ��ġ�����԰�һ��Ctrl+Shift+q�����ϼ���ǿ���˳����ԣ�Goodluck �ף�")

HotKeySet("^+q", "Terminate")

If $answer = 1 Then
   TestMain()
   MsgBox(0, "�������", "�󹦸��")
   Exit
EndIf


Func ReadAndInputFromFile($file)
   Local $iCount = 1
   Local $iFileE = 1
   ;ѭ����ȡ�ļ�ÿһ�У��������������
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
   ;���ý���ΪPS2�����
     Sleep(1000)
     MouseClick("left")
    
   ;�������ļ�
   Local $file = FileOpen(@DesktopDir & "\1.txt", 0)
   If $file = -1 Then
       MsgBox(0, "������", "���Ƕ������ǲ��ǰ�  1.txt  �������Ūû�ˣ���ס������һ�Ҫ����������~")
       Exit
   EndIf

   ReadAndInputFromFile($file)
  
   ;�ر��ļ�����
   FileClose($file)
EndFunc

Func Terminate()
    Exit 0
EndFunc