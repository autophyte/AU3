#include <Excel.au3>
#include <array.au3>

;---------------------------------------------------------------------------
;下面定义变量
;---------------------------------------------------------------------------

Global $hExcel               ;定义EXCEL文件的句柄
Global $hFileSB
Global $hFileSBE
Global $hFileReport

Global $sRange               ;需要读取的CELL的区域
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

; --------------------<原始数据>--------------------------------
Global $oItemID[$oArrayLen]     ;物品道具ID --------------RnC1
;物品名称，武器RnC1，道具RnC4，插件RnC3
Global $oItemName[$oArrayLen]
;物品SC价格，武器RnC6，道具RnC6，插件RnC7
Global $oItemSCPrice[$oArrayLen]
;物品CP价格，武器RnC7
Global $oItemCPPrice[$oArrayLen]
;物品所属阵营，武器RnC11，道具RnC8
Global $oItemFaction[$oArrayLen]
;物品所属职业载具，武器RnC12，道具RnC9
Global $oItemClass[$oArrayLen]
;道具是否上架，武器RnC10，道具RnC7，插件RnC5
Global $oItemSale[$oArrayLen]

Global $sbMBID[$sbArrayLen]
Global $sbSCPrice[$sbArrayLen]
Global $sbCPPrice[$sbArrayLen]
Global $sbSTime[$sbArrayLen]
Global $sbETime[$sbArrayLen]
Global $sbStatus[$sbArrayLen]

; 将临时数据添加到原始数据数组中
Func PushToOriginalArray()
   If $tmpSale = "是" Then
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

; 打开存储价格的EXCEL文件
Func OpenExcel()
   $hExcel = _ExcelBookOpen(@DesktopDir & "\TestData\1.xlsx", 0, True)
   If @error = 1 OR @error = 2 Then
       SetErrorInformation(1)
       MsgBox(0, "Error!", "Unable to Open the Excel Object")
       Exit
   EndIf

   ; 激活第1/2个sheet，读取一个武器数据到临时数据中--原始数据 sheet1, sheet2
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

   ; 激活第3个sheet，读取一个装备数据到临时数据中--原始数据 sheet3
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

   ; 激活第3个sheet，读取一个插件数据到临时数据中--原始数据 sheet4
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
          $sbSTime[$sbArrayItr]               = $arrayString[12] ;从原来11列到现在12列
          $sbETime[$sbArrayItr]               = $arrayString[13] ;从原来12列到现在13列
          $sbStatus[$sbArrayItr]               = $arrayString[15] ;从原来14列到现在15列
		  
		  ; SC点原来是第8列表示，现在由8、9共同，8为类型，9为数值
		  If $arrayString[10] = 7000 Then
			$sbSCPrice[$sbArrayItr]          = $arrayString[11]
		Else
			$sbSCPrice[$sbArrayItr]          = 0
		EndIf
		  
			; CP点原来是第9、10列表示，更新为10、11列
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
   MakeErrInf("PACKAGE/ITEM ID", "ERROR_TYPE", "游戏中SC价格", "需求SC价格", "游戏中CP价格", "需求中CP价格", "道具名称")
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

; "ERROR_TYPE_0" "道具包信息错误"
; "ERROR_TYPE_1" "道具只在游戏中有，需求文件中找不到"
; "ERROR_TYPE_2" "游戏中道具SC价格与需求文件中价格不一致"
; "ERROR_TYPE_3" "游戏中道具CP价格与需求文件中价格不一致"
; "ERROR_TYPE_4" "游戏中道具CP和SC价格均与需求文件中价格不一致"
; "ERROR_TYPE_5" "游戏中上架状态与需求文件中不一致"
; "PACKAGE/ITEM ID"  "ERROR_TYPE"  "游戏中SC价格"  "需求SC价格"  "游戏中CP价格"  "需求中CP价格"  "游戏上架状态"  "需求上架" "道具名称/ID"
Func MakeErrInf($srcID, $errType, $ssc, $osc, $scp, $ocp, $itmName)
   Local $tmpLine = $srcID & @TAB & $errType &  @TAB & $ssc &  @TAB & $osc &  @TAB & $scp &  @TAB & $ocp &  @TAB & $itmName
   FileWriteLine($hFileReport, $tmpLine)
EndFunc

Func CompareDatas()
   Local $retVal = GetSBEFromByMBID($sbeMBID)
   If $retVal = -1 Then
       MakeErrInf($sbeMBID, "ERROR_TYPE_0", 0, 0, 0, 0, "")
   ElseIf $retVal = -2 Then
       ;return -2表示status为0，不上架
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