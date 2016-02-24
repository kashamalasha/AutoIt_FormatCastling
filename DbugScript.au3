#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <Dbug.au3>

Execute(Dbug(5))
Global $sFormat
Execute(Dbug(6))
Global $sTNS = 'osp';$CmdLine[1]
Execute(Dbug(7))
Global $sUSERNAME = 'sdd';$CmdLine[2]
Execute(Dbug(8))
Global $sPASSWORD = 'sdd';$CmdLine[3]

Execute(Dbug(10))
Func WhoIsThere()

Execute(Dbug(12))
    Local $oSQL = ObjCreate("ADODB.Connection")
Execute(Dbug(13))
    With $oSQL
Execute(Dbug(14))
	  .ConnectionString =("Provider='OraOLEDB.Oracle';Data Source=" & $sTNS & ";User Id=" & $sUSERNAME & ";Password=" & $sPASSWORD & ";")
Execute(Dbug(15))
	  .Open
Execute(Dbug(16))
    EndWith
Execute(Dbug(17))
    Local $oSQLrs = ObjCreate("ADODB.RecordSet")
Execute(Dbug(18))
    With $oSQLrs
Execute(Dbug(19))
	   .ActiveConnection = $oSQL
Execute(Dbug(20))
	   .Source = "SELECT EXT_STRING " & _
		 	     "FROM SDD.DEPARTMENT_EXT " & _
		 	     "WHERE EXT_NAME = 'FORMAT' " & _
		 	     "AND ID_DEPARTMENT = ( " & _
		 	     "  SELECT ID_DEPARTMENT " & _
		 	     "  FROM SDD.DEPARTMENT " & _
		 	     "  WHERE IS_HOST = 1)"
Execute(Dbug(27))
	   .Open
Execute(Dbug(28))
    EndWith
Execute(Dbug(29))
    $sFormat = $oSQLrs.Fields(0).Value
Execute(Dbug(30))
    $oSQL.Close

Execute(Dbug(32))
EndFunc ;==>WhoIsThere

Execute(Dbug(34))
Func TheCastling()

Execute(Dbug(36))
	Local $oSQL = ObjCreate("ADODB.Connection")
Execute(Dbug(37))
    With $oSQL
Execute(Dbug(38))
	  .ConnectionString =("Provider='OraOLEDB.Oracle';Data Source=" & $sTNS & ";User Id=" & $sUSERNAME & ";Password=" & $sPASSWORD & ";")
Execute(Dbug(39))
	  .Open
Execute(Dbug(40))
    EndWith

Execute(Dbug(42))
    If $sFormat = 'SUPER' Then
Execute(Dbug(43))
        $oSQL.Execute(  "UPDATE SDD.DEPARTMENT_EXT DE " & _
						"SET DE.EXT_STRING = 'GIPER' " & _
						"WHERE DE.EXT_NAME = 'FORMAT' AND " & _
						"DE.ID_DEPARTMENT = ( " & _
						   "SELECT D.ID_DEPARTMENT " & _
						   "FROM SDD.DEPARTMENT D " & _
						   "WHERE D.IS_HOST = 1)" _
			   )
Execute(Dbug(51))
    Else
Execute(Dbug(52))
        $oSQL.Execute(  "UPDATE SDD.DEPARTMENT_EXT DE " & _
						"SET DE.EXT_STRING = 'SUPER' " & _
						"WHERE DE.EXT_NAME = 'FORMAT' AND " & _
						"DE.ID_DEPARTMENT = ( " & _
						   "SELECT D.ID_DEPARTMENT " & _
						   "FROM SDD.DEPARTMENT D " & _
						   "WHERE D.IS_HOST = 1)" _
			   )
Execute(Dbug(60))
    EndIf
Execute(Dbug(61))
    $oSQL.Close
Execute(Dbug(62))
	WhoIsThere()

Execute(Dbug(64))
EndFunc ;==>TheCastling

Execute(Dbug(66))
Func Main()
Execute(Dbug(67))
	Local $bSet = False
Execute(Dbug(68))
    Local $iCount = 0
Execute(Dbug(69))
    Local $sSetFormat

Execute(Dbug(71))
    GUICreate("TC X-K?", 120, 120)
Execute(Dbug(72))
    $idChange = GUICtrlCreateButton("Format Pic", 10, 10, 100, 100, $BS_ICON)

Execute(Dbug(74))
    WhoIsThere()
Execute(Dbug(75))
    $sSetFormat = $sFormat
Execute(Dbug(76))
    GUISetState(@SW_SHOW)

Execute(Dbug(78))
    While 1
        Switch GUIGetMsg()
             Case $GUI_EVENT_CLOSE
Execute(Dbug(81))
                ExitLoop
Execute(Dbug(82))
			 Case $idChange
Execute(Dbug(83))
                TheCastling()
Execute(Dbug(84))
				$bSet = False
Execute(Dbug(85))
        EndSwitch
Execute(Dbug(86))
		If $bSet = False Then
Execute(Dbug(87))
		   If $sFormat = 'SUPER' Then
Execute(Dbug(88))
			   GUICtrlSetImage($idChange, @ScriptDir & "\super.ico")
Execute(Dbug(89))
			   $bSet = True
Execute(Dbug(90))
		   Else
Execute(Dbug(91))
			   GUICtrlSetImage($idChange, @ScriptDir & "\giper.ico")
Execute(Dbug(92))
			   $bSet = True
Execute(Dbug(93))
		   EndIf
Execute(Dbug(94))
		EndIf
Execute(Dbug(95))
		If $iCount = 1000 Then
Execute(Dbug(96))
		     WhoIsThere()
Execute(Dbug(97))
		     If StringCompare($sFormat, $sSetFormat) <> 0 Then
Execute(Dbug(98))
			    $bSet = False
Execute(Dbug(99))
			 EndIf
Execute(Dbug(100))
			 $iCount = 0
Execute(Dbug(101))
		Else
Execute(Dbug(102))
			 $iCount += 1
Execute(Dbug(103))
		EndIf
Execute(Dbug(104))
    WEnd
Execute(Dbug(105))
EndFunc   ;==>Main

 ;Run the app
Execute(Dbug(108))
 Main()
