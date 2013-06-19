#include <Excel.au3>
#include <array.au3>

;---------------------------------------------------------------------------
;���涨�����
;---------------------------------------------------------------------------

Global $hExcel               ;����EXCEL�ļ��ľ��
Global $hFileSB
Global $hFileSBE
Global $hFileReport

Global $sRange               ;��Ҫ��ȡ��CELL������
Global $oArrayLen               = 4000
Global $sbArrayLen               = 4000
Global $oArrayItr               = 0
Global $sbArrayItr               = 0

Global $sbeMBID
Global $sbeItemID
Global $sbeSCPrice
Global $sbeCPPrice

Global $tmpItemID
Global $tmpName
Global $tmpSCPrice
Global $tmpCPPrice
Global $tmpFact
Global $tmpClass
Global $tmpSale

; --------------------<ԭʼ����>--------------------------------
Global $oItemID[$oArrayLen]     ;��Ʒ����ID --------------RnC1
;��Ʒ���ƣ�����RnC1������RnC4�����RnC3
Global $oItemName[$oArrayLen]
;��ƷSC�۸�����RnC6������RnC6�����RnC7
Global $oItemSCPrice[$oArrayLen]
;��ƷCP�۸�����RnC7
Global $oItemCPPrice[$oArrayLen]
;��Ʒ������Ӫ������RnC11������RnC8
Global $oItemFaction[$oArrayLen]
;��Ʒ����ְҵ�ؾߣ�����RnC12������RnC9
Global $oItemClass[$oArrayLen]
;�����Ƿ��ϼܣ�����RnC10������RnC7�����RnC5
Global $oItemSale[$oArrayLen]

Global $sbMBID[$sbArrayLen]
Global $sbSCPrice[$sbArrayLen]
Global $sbCPPrice[$sbArrayLen]
Global $sbSTime[$sbArrayLen]
Global $sbETime[$sbArrayLen]
Global $sbStatus[$sbArrayLen]

; ����ʱ������ӵ�ԭʼ����������
Func PushToOriginalArray()
   If $tmpSale = "��" Then
       $tmpSale = 1
   Else
       $tmpSale = 0
   EndIf
  
     If $oArrayItr < $sbArrayLen Then
          $oItemID[$oArrayItr]          = $tmpItemID
          $oItemName[$oArrayItr]          = $tmpName
          $oItemSCPrice[$oArrayItr]     = $tmpSCPrice
          $oItemCPPrice[$oArrayItr]     = $tmpCPPrice
          $oItemFaction[$oArrayItr]     = $tmpFact
          $oItemClass[$oArrayItr]          = $tmpClass
          $oItemSale[$oArrayItr]          = $tmpSale
     Else
          $oArrayLen = _ArrayAdd($oItemID[$oArrayItr], $tmpItemID)
          _ArrayAdd($oItemName[$oArrayItr], $tmpName)
          _ArrayAdd($oItemSCPrice[$oArrayItr], $tmpSCPrice)
          _ArrayAdd($oItemCPPrice[$oArrayItr], $tmpCPPrice)
          _ArrayAdd($oItemFaction[$oArrayItr], $tmpFact)
          _ArrayAdd($oItemClass[$oArrayItr], $tmpClass)
          _ArrayAdd($oItemSale[$oArrayItr], $tmpSale)
     EndIf
     $oArrayItr = $oArrayItr + 1
EndFunc

Func GetFromOriginalArray($sourceItemID)
   Local $tmpIndex = _ArraySearch($oItemID, $sourceItemID)
   If $tmpIndex = -1 Then
       $tmpItemID     = 0
       $tmpName          = ""
       $tmpSCPrice     = 0
       $tmpCPPrice     = 0
       $tmpFact          = 0
       $tmpClass          = 0
       $tmpSale          = 0
       return -1
   Else
       $tmpItemID     = $oItemID[$tmpIndex]
       $tmpName          = $oItemName[$tmpIndex]
       $tmpSCPrice     = $oItemSCPrice[$tmpIndex]
       $tmpCPPrice     = $oItemCPPrice[$tmpIndex]
       $tmpFact          = $oItemFaction[$tmpIndex]
       $tmpClass          = $oItemClass[$tmpIndex]
       $tmpSale          = $oItemSale[$tmpIndex]
       return 0
   EndIf
EndFunc

; �򿪴洢�۸��EXCEL�ļ�
Func OpenExcel()
   $hExcel = _ExcelBookOpen(@DesktopDir & "\TestData\1.xlsx", 0, True)
   If @error = 1 OR @error = 2 Then
       SetErrorInformation(1)
       MsgBox(0, "Error!", "Unable to Open the Excel Object")
       Exit
   EndIf

   ; �����1/2��sheet����ȡһ���������ݵ���ʱ������--ԭʼ���� sheet1, sheet2
   _ExcelSheetActivate($hExcel, 1)
   Local $aArray0 = _ExcelReadSheetToArray($hExcel, 2)
   For $iRow = 1 To $aArray0[0][0] Step 1
       IF $aArray0[0][1] >= 12 Then
          $tmpItemID          = $aArray0[$iRow][1]
          $tmpName          = $aArray0[$iRow][3]
          $tmpSCPrice     = $aArray0[$iRow][6]
          $tmpCPPrice     = $aArray0[$iRow][7]
          $tmpSale          = $aArray0[$iRow][10]
          $tmpFact          = $aArray0[$iRow][11]
          $tmpClass          = $aArray0[$iRow][12]
          PushToOriginalArray()
       EndIf
   Next
   _ExcelSheetActivate($hExcel, 2)
   Local $aArray1 = _ExcelReadSheetToArray($hExcel, 2)
   For $iRow = 1 To $aArray1[0][0] Step 1
       IF $aArray1[0][1] >= 12 Then
          $tmpItemID          = $aArray1[$iRow][1]
          $tmpName          = $aArray1[$iRow][3]
          $tmpSCPrice     = $aArray1[$iRow][6]
          $tmpCPPrice     = $aArray1[$iRow][7]
          $tmpSale          = $aArray1[$iRow][10]
          $tmpFact          = $aArray1[$iRow][11]
          $tmpClass          = $aArray1[$iRow][12]
          PushToOriginalArray()
       EndIf
   Next

   ; �����3��sheet����ȡһ��װ�����ݵ���ʱ������--ԭʼ���� sheet3
   _ExcelSheetActivate($hExcel, 3)
   Local $aArray2 = _ExcelReadSheetToArray($hExcel, 2)
   For $iRow = 1 To $aArray2[0][0] Step 1
       IF $aArray2[0][1] >= 9 Then
          $tmpItemID          = $aArray2[$iRow][1]
          $tmpName          = $aArray2[$iRow][3]
          $tmpSCPrice     = $aArray2[$iRow][6]
          $tmpCPPrice     = 0
          $tmpSale          = $aArray2[$iRow][7]
          $tmpFact          = $aArray2[$iRow][8]
          $tmpClass          = $aArray2[$iRow][9]
          PushToOriginalArray()
       EndIf
   Next

   ; �����3��sheet����ȡһ��������ݵ���ʱ������--ԭʼ���� sheet4
   _ExcelSheetActivate($hExcel, 4)
   Local $aArray3 = _ExcelReadSheetToArray($hExcel, 2)
   For $iRow = 1 To $aArray3[0][0] Step 1
       IF $aArray3[0][1] >= 7 Then
          $tmpItemID          = $aArray3[$iRow][1]
          $tmpName          = $aArray3[$iRow][4]
          $tmpSale          = $aArray3[$iRow][5]
          $tmpSCPrice     = $aArray3[$iRow][7]
          $tmpCPPrice     = 0
          $tmpFact          = 0
          $tmpClass          = 0
          PushToOriginalArray()
       EndIf
   Next
  
   _ExcelBookClose($hExcel)
EndFunc

Func SetErrorInformation($iCode)
EndFunc

Func PushToSBArray($arrayString)
     If $sbArrayItr < $oArrayLen Then
          $sbMBID[$sbArrayItr]               = $arrayString[1]
          $sbSTime[$sbArrayItr]               = $arrayString[12] ;��ԭ��11�е�����12��
          $sbETime[$sbArrayItr]               = $arrayString[13] ;��ԭ��12�е�����13��
          $sbStatus[$sbArrayItr]               = $arrayString[15] ;��ԭ��14�е�����15��
		  
		  ; SC��ԭ���ǵ�8�б�ʾ��������8��9��ͬ��8Ϊ���ͣ�9Ϊ��ֵ
		  If $arrayString[10] = 7000 Then
			$sbSCPrice[$sbArrayItr]          = $arrayString[11]
		Else
			$sbSCPrice[$sbArrayItr]          = 0
		EndIf
		  
			; CP��ԭ���ǵ�9��10�б�ʾ������Ϊ10��11��
          If $arrayString[9] = 10 Then
               $sbCPPrice[$sbArrayItr]          = $arrayString[10]
          Else
               $sbCPPrice[$sbArrayItr]          = 0
          EndIf
       Else
          _ArrayAdd($sbMBID[$sbArrayItr], $arrayString[1])
          _ArrayAdd($sbSCPrice[$sbArrayItr], $arrayString[8])
          _ArrayAdd($sbSTime[$sbArrayItr], $arrayString[11])
          _ArrayAdd($sbETime[$sbArrayItr], $arrayString[12])
          _ArrayAdd($sbStatus[$sbArrayItr], $arrayString[14])
          If $arrayString[9] = 10 Then
               _ArrayAdd($sbCPPrice[$sbArrayItr], $arrayString[10])
          Else
               _ArrayAdd($sbCPPrice[$sbArrayItr], 0)
          EndIf         
     EndIf
     $sbArrayItr = $sbArrayItr + 1
EndFunc

Func OpenReportFile()
   $hFileReport = FileOpen(@DesktopDir & "\TestData\Report.txt", 9)
   If $hFileReport = -1 Then
       SetErrorInformation(1)
       MsgBox(0, "Error", "Unable to open file.")
       Exit
   EndIf
   Local $strLine = @YEAR & "-" & @MON & "-" & @MDAY & "-" & @HOUR & "-" & @MIN & "-" & @SEC
   FileWriteLine($hFileReport, ">>>>>>>>>>>>>>>" & $strLine & "<<<<<<<<<<<<<<<")
   MakeErrInf("PACKAGE/ITEM ID", "ERROR_TYPE", "��Ϸ��SC�۸�", "����SC�۸�", "��Ϸ��CP�۸�", "������CP�۸�", "��������")
EndFunc

Func OpenSBFile()
   $hFileSB = FileOpen(@DesktopDir & "\TestData\StoreBundles.txt", 0)
   If $hFileSB = -1 Then
       SetErrorInformation(1)
       MsgBox(0, "Error", "Unable to open file.")
       Exit
   EndIf

   Local $line = FileReadLine($hFileSB)
   If @error = -1 Then
       SetErrorInformation(1)
       Exit
   EndIf
  
   While 1
       $line = FileReadLine($hFileSB)
       If @error = -1 Then
          ExitLoop
       EndIf
      
       Local $arrayString = StringSplit($line, '^', 1)
       If @error <> 1 AND $arrayString[0] > 30 Then
          PushToSBArray($arrayString)
       EndIf
   WEnd
   FileClose($hFileSB)
EndFunc

; "ERROR_TYPE_0" "���߰���Ϣ����"
; "ERROR_TYPE_1" "����ֻ����Ϸ���У������ļ����Ҳ���"
; "ERROR_TYPE_2" "��Ϸ�е���SC�۸��������ļ��м۸�һ��"
; "ERROR_TYPE_3" "��Ϸ�е���CP�۸��������ļ��м۸�һ��"
; "ERROR_TYPE_4" "��Ϸ�е���CP��SC�۸���������ļ��м۸�һ��"
; "ERROR_TYPE_5" "��Ϸ���ϼ�״̬�������ļ��в�һ��"
; "PACKAGE/ITEM ID"  "ERROR_TYPE"  "��Ϸ��SC�۸�"  "����SC�۸�"  "��Ϸ��CP�۸�"  "������CP�۸�"  "��Ϸ�ϼ�״̬"  "�����ϼ�" "��������/ID"
Func MakeErrInf($srcID, $errType, $ssc, $osc, $scp, $ocp, $itmName)
   Local $tmpLine = $srcID & @TAB & $errType &  @TAB & $ssc &  @TAB & $osc &  @TAB & $scp &  @TAB & $ocp &  @TAB & $itmName
   FileWriteLine($hFileReport, $tmpLine)
EndFunc

Func CompareDatas()
   Local $retVal = GetSBEFromByMBID($sbeMBID)
   If $retVal = -1 Then
       MakeErrInf($sbeMBID, "ERROR_TYPE_0", 0, 0, 0, 0, "")
   ElseIf $retVal = -2 Then
       ;return -2��ʾstatusΪ0�����ϼ�
   Else
       $retVal = GetFromOriginalArray($sbeItemID)
       If $retVal = -1 Then
          MakeErrInf($sbeItemID, "ERROR_TYPE_1", 0, 0, 0, 0, "")
       Else
          If $sbeSCPrice <> $tmpSCPrice Then
               If $sbeCPPrice <> $tmpCPPrice Then
                  MakeErrInf($sbeItemID, "ERROR_TYPE_4", $sbeSCPrice, $tmpSCPrice, $sbeCPPrice, $tmpCPPrice, $tmpName)
               Else
                  MakeErrInf($sbeItemID, "ERROR_TYPE_2", $sbeSCPrice, $tmpSCPrice, $sbeCPPrice, $tmpCPPrice, $tmpName)
               EndIf
          Else
               If $sbeCPPrice <> $tmpCPPrice Then
                  MakeErrInf($sbeItemID, "ERROR_TYPE_3", $sbeSCPrice, $tmpSCPrice, $sbeCPPrice, $tmpCPPrice, $tmpName)
               EndIf
          EndIf
       EndIf
   EndIf
EndFunc

Func GetSBEFromByMBID($sourceMBID)
   Local $tmpIndex = _ArraySearch($sbMBID, $sourceMBID)
   If $tmpIndex = -1 Then
       $sbeSCPrice = 0
       $sbeCPPrice = 0
       return -1
   Else
       $sbeSCPrice = $sbSCPrice[$tmpIndex]
       $sbeCPPrice = $sbCPPrice[$tmpIndex]
   EndIf
   If $sbStatus[$tmpIndex] = 0 Then
       Return -2
   EndIf
  
   return 0
EndFunc

Func OpenSBEFile()
   $hFileSBE = FileOpen(@DesktopDir & "\TestData\StoreBundleEntries.txt", 0)
   If $hFileSBE = -1 Then
       SetErrorInformation(1)
       MsgBox(0, "Error", "Unable to open file.")
       Exit
   EndIf

   Local $line = FileReadLine($hFileSBE)
   If @error = -1 Then
       SetErrorInformation(1)
       Exit
   EndIf
  
   While 1
       Local $line = FileReadLine($hFileSBE)
       If @error = -1 Then
          ExitLoop
       EndIf
      
       Local $arrayString = StringSplit($line, '^', 1)
       If @error <> 1 AND $arrayString[0] > 4 Then
          ;marketing_bundle_id is $arrayString[1]
          ;game item id is $arrayString[4]
          $sbeMBID          = $arrayString[1]
          $sbeItemID          = $arrayString[4]
          CompareDatas()
       EndIf
   WEnd
EndFunc

Func MainTestFun()
   OpenExcel()
   OpenReportFile()
   OpenSBFile()
   OpenSBEFile()
   FileClose($hFileReport)
EndFunc


MainTestFun()