#AutoIt3Wrapper_Icon=FC.ico

#include <ButtonConstants.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
;#include <Dbug.au3>

Global $sFormat

If $CmdLine[0] < 3 Then
   MsgBox(16, @ScriptName & ": Ошибка", "Неверно указана строка соединения с БД" & _
			@CRLF & "Укажи параметры следующим образом:" & _
			@CRLF & @CRLF & """FormatCastling.exe TNS USER PASSWORD""")
   Exit
EndIf

Global $sTNS = $CmdLine[1]
Global $sUsername = $CmdLine[2]
Global $sPassword = $CmdLine[3]
Global $sCurDir

If @Compiled Then
	FileInstall("giper.ico", @TempDir & "\giper.ico")
	FileInstall("super.ico", @TempDir & "\super.ico")
    FileChangeDir(@TempDir)
EndIf

Func WhoIsThere()
    Local $oSQL = ObjCreate("ADODB.Connection")

    With $oSQL
      .ConnectionString =("Provider='OraOLEDB.Oracle';" & _
                          "Data Source=" & $sTNS & ";" & _
                          "User Id=" & $sUsername & ";" & _
                          "Password=" & $sPassword & ";")
      .Open
    EndWith
    Local $oSQLrs = ObjCreate("ADODB.RecordSet")
    With $oSQLrs
      .ActiveConnection = $oSQL
      .Source = "SELECT EXT_STRING " & _
                "FROM SDD.DEPARTMENT_EXT " & _
                "WHERE EXT_NAME = 'FORMAT' " & _
                "AND ID_DEPARTMENT = ( " & _
                "  SELECT ID_DEPARTMENT " & _
                "  FROM SDD.DEPARTMENT " & _
                "  WHERE IS_HOST = 1)"
      .Open
    EndWith
    $sFormat = $oSQLrs.Fields(0).Value
    $oSQL.Close

EndFunc ;==>WhoIsThere

Func TheCastling()
    Local $oSQL = ObjCreate("ADODB.Connection")

    With $oSQL
      .ConnectionString =("Provider='OraOLEDB.Oracle';" & _
                          "Data Source=" & $sTNS & ";" & _
                          "User Id=" & $sUsername & ";" & _
                          "Password=" & $sPassword & ";")
      .Open
    EndWith

    If $sFormat = 'SUPER' Then
        $oSQL.Execute("UPDATE SDD.DEPARTMENT_EXT DE " & _
                      "SET DE.EXT_STRING = 'GIPER' " & _
                      "WHERE DE.EXT_NAME = 'FORMAT' AND " & _
                      "DE.ID_DEPARTMENT = ( " & _
                      "   SELECT D.ID_DEPARTMENT " & _
                      "   FROM SDD.DEPARTMENT D " & _
                      "   WHERE D.IS_HOST = 1)" _
               )
    Else
        $oSQL.Execute("UPDATE SDD.DEPARTMENT_EXT DE " & _
                      "SET DE.EXT_STRING = 'SUPER' " & _
                      "WHERE DE.EXT_NAME = 'FORMAT' AND " & _
                      "DE.ID_DEPARTMENT = ( " & _
                      "   SELECT D.ID_DEPARTMENT " & _
                      "   FROM SDD.DEPARTMENT D " & _
                      "   WHERE D.IS_HOST = 1)" _
               )
    EndIf
    $oSQL.Close
    WhoIsThere()

EndFunc ;==>TheCastling

Func Main()
    Local $bSet = False
    Local $iCount = TimerInit()
    Local $sSetFormat

    GUICreate("FC", 120, 120, @DesktopWidth - 160, 100, Default, $WS_EX_TOPMOST)
    GUISetIcon(@ScriptDir & "\FC.ico")
	$idChange = GUICtrlCreateButton("Format Pic", 10, 10, 100, 100, $BS_ICON)

    WhoIsThere()
    $sSetFormat = $sFormat
    GUISetState(@SW_SHOW)

    While 1
        Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				If @Compiled Then
					FileDelete(@TempDir & "/giper.ico")
					FileDelete(@TempDir & "/super.ico")
				EndIf
                ExitLoop
            Case $idChange
                TheCastling()
                $bSet = False
        EndSwitch
        If $bSet = False Then
            If $sFormat = 'SUPER' Then
                GUICtrlSetImage($idChange, @WorkingDir & "\super.ico")
                $bSet = True
            Else
                GUICtrlSetImage($idChange, @WorkingDir & "\giper.ico")
                $bSet = True
			 EndIf
			 GUICtrlSetData($idChange, $sFormat)
             $sSetFormat = $sFormat
        EndIf
        If TimerDiff($iCount) >= 1000 Then
            WhoIsThere()
            If StringCompare($sFormat, $sSetFormat) <> 0 Then
                $bSet = False
            EndIf
                $iCount = TimerInit()
        EndIf
    WEnd
EndFunc   ;==>Main

 ;Run the app
 Main()