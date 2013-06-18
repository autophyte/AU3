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
   ;�趨��ͼ��Χ��������û�б����������ֵ�����ʱ�������ͼ�����ں����ж�
   MsgBox(0x40000, "�趨��ͼλ��", "������Ƶ���ͼ�������Ͻǣ�Ȼ�󰴿ո��")
   $cap_left_top = MouseGetPos()
   MsgBox(0x40000, "�趨��ͼλ��", "������Ƶ���ͼ�������½ǣ�Ȼ�󰴿ո��")
   $cap_right_bootom = MouseGetPos()

   ;�趨��ɰ�ť�����򣬽ű��á���ɰ�ť����ɫ��ȷ�������Ƿ����
   ;����������ɫ��Χ��λ��1Ϊ��ɰ�ť�����Ͻǣ�λ��2Ϊ��ɰ�ť�����½ǣ����Ը�����
   MsgBox(0x40000, "�趨��ťλ��", "������Ƶ���ɰ�ť���Ͻǣ�Ȼ�󰴿ո��")
   $pos_left_top = MouseGetPos()
   MsgBox(0x40000, "�趨��ťλ��", "������Ƶ���ɰ�ť���½ǣ�Ȼ�󰴿ո��")
   $pos_right_bootom = MouseGetPos()
 
   ;�趨������λ��
   MsgBox(0x40000, "�趨�����λ��", "�뽫����Ƶ������λ�ã����¿ո��")
   $pos_input = MouseGetPos()
EndFunc


Func ReadAndInputFromFile($file, $file_r)
 
   Local $iLine = 1
   ;ѭ����ȡ�ļ�ÿһ�У��������������
   while 1
       $strLine = FileReadLine($file)
       If @error = -1 Then
          ExitLoop
       Endif
       
        If StringLen($strLine) <= 1 Then
            $strLine = $strLine & " ";
        EndIf

       ;����������������ڣ���������������
       MouseClick("left", $pos_input[0], $pos_input[1])
       Send(("^{a}"))
       Sleep(100)
       Send("{BS}")
     
       ;�����ȡ��������?
       Send($strLine)
        Send("{ENTER}")
       Sleep(2000)
     
       ;������ɫ2780482�������ɫ����ɫ������ɰ�ť������Ҳ�������ֿ���ʱ����Χ��Ӧ�ð��������ɫ
       ;�����ҵ������ɫ����ʱ��ʾ����û�б������������б��У�Ӧ�ü�¼
       Local $coord = PixelSearch($pos_left_top[0], $pos_left_top[1], $pos_right_bootom[0], $pos_right_bootom[1], 2780482, 5)
      If Not @error Then
          ;��ͼ
          _ScreenCapture_Capture(@MyDocumentsDir & "\CharacterNameImage_" & $iLine & ".jpg", $cap_left_top[0], $cap_left_top[1], $cap_right_bootom[0], $cap_right_bootom[1])

          ;��¼���ļ���
          FileWriteLine($file_r, $iLine & @TAB & $strLine & @TAB & "\CharacterNameImage_" & $iLine & ".jpg" & @CRLF)
          FileFlush($file_r)
       EndIf
     
       $iLine = $iLine + 1
   WEnd
EndFunc

Func TestMain()
   SetTestRange()
 
   ;�������ļ�
   Local $file = FileOpen(@DesktopDir & "\1.txt", 0)
   If $file = -1 Then
       MsgBox(0, "������", "���Ƕ������ǲ��ǰ�  1.txt  �������Ūû�ˣ���ס������һ�Ҫ����������~")
       Exit
   EndIf
 
   ;��report�ļ�
   Local $file_r = FileOpen(@MyDocumentsDir & "\CharacterNameReport.txt", 10)
   If $file = -1 Then
       MsgBox(0, "������", "���Ƕ�������ԭ����Ҳ��֪��������������û�д����ɹ�~")
       Exit
   EndIf
   FileWriteLine($file_r, "LineNO." & @TAB & "Name" & @TAB & "Picture" & @CRLF)

   ReadAndInputFromFile($file, $file_r)
 
   FileClose($file)
   FileClose($file_r)
EndFunc


Func Terminate()
   MsgBox(0, "�����ж�", "��ǰ���Խ��ȣ�" & $strLine)
   Exit 0
EndFunc

TestMain()
MsgBox(0, "�������", "�󹦸��")