#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <Dbug.au3>

Execute(Dbug(5))
Global $sFormat
Execute(Dbug(6))
Global Const $sTNS = $CmdLine[1]
Execute(Dbug(7))
Global Const $sUSERNAME = $CmdLine[2]
Execute(Dbug(8))
Global Const $sPASSWORD = $CmdLine[3]

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
Execute(Dbug(31))
 EndFunc ;==>WhoIsThere

Execute(Dbug(33))
 Func TheCastling()

Execute(Dbug(35))
	Local $oSQL = ObjCreate("ADODB.Connection")
Execute(Dbug(36))
    With $oSQL
Execute(Dbug(37))
	  .ConnectionString =("Provider='OraOLEDB.Oracle';Data Source=" & $sTNS & ";User Id=" & $sUSERNAME & ";Password=" & $sPASSWORD & ";")
Execute(Dbug(38))
	  .Open
Execute(Dbug(39))
    EndWith

Execute(Dbug(41))
    If $sFormat = 'SUPER' Then
Execute(Dbug(42))
        $oSQL.Execute(  "UPDATE SDD.DEPARTMENT_EXT DE " & _
						"SET DE.EXT_STRING = 'GIPER' " & _
						"WHERE DE.EXT_NAME = 'FORMAT' AND " & _
						"DE.ID_DEPARTMENT = ( " & _
						   "SELECT D.ID_DEPARTMENT " & _
						   "FROM SDD.DEPARTMENT D " & _
						   "WHERE D.IS_HOST = 1)" _
			   )
Execute(Dbug(50))
    Else
Execute(Dbug(51))
        $oSQL.Execute(  "UPDATE SDD.DEPARTMENT_EXT DE " & _
						"SET DE.EXT_STRING = 'SUPER' " & _
						"WHERE DE.EXT_NAME = 'FORMAT' AND " & _
						"DE.ID_DEPARTMENT = ( " & _
						   "SELECT D.ID_DEPARTMENT " & _
						   "FROM SDD.DEPARTMENT D " & _
						   "WHERE D.IS_HOST = 1)" _
			   )
Execute(Dbug(59))
    EndIf
Execute(Dbug(60))
    $oSQL.Close
Execute(Dbug(61))
	WhoIsThere()
Execute(Dbug(62))
 EndFunc ;==>TheCastling

Execute(Dbug(64))
 Func Main()
Execute(Dbug(65))
	Local $bSet = False

Execute(Dbug(67))
    GUICreate("TC X-K?", 120, 120)
Execute(Dbug(68))
    $idChange = GUICtrlCreateButton("Format Pic", 10, 10, 100, 100, $BS_ICON)

Execute(Dbug(70))
    WhoIsThere()

Execute(Dbug(72))
    GUISetState(@SW_SHOW)

Execute(Dbug(74))
    While 1
        Switch GUIGetMsg()
             Case $GUI_EVENT_CLOSE
Execute(Dbug(77))
                ExitLoop

Execute(Dbug(79))
			 Case $idChange
Execute(Dbug(80))
                TheCastling()
Execute(Dbug(81))
				$bSet = False

Execute(Dbug(83))
        EndSwitch

Execute(Dbug(85))
		If $bSet = False Then
Execute(Dbug(86))
		   If $sFormat = 'SUPER' Then
Execute(Dbug(87))
			   GUICtrlSetImage($idChange, @ScriptDir & "\super.ico")
Execute(Dbug(88))
			   $bSet = True
Execute(Dbug(89))
		   Else
Execute(Dbug(90))
			   GUICtrlSetImage($idChange, @ScriptDir & "\giper.ico")
Execute(Dbug(91))
			   $bSet = True
Execute(Dbug(92))
		   EndIf
Execute(Dbug(93))
		EndIf

Execute(Dbug(95))
    WEnd
Execute(Dbug(96))
 EndFunc   ;==>Main

 ;Run the app
Execute(Dbug(99))
 Main()
