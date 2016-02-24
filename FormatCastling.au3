#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
;#include <Dbug.au3>

Global $sFormat
Global $sTNS = $CmdLine[1]
Global $sUSERNAME = $CmdLine[2]
Global $sPASSWORD = $CmdLine[3]

Func WhoIsThere()

    Local $oSQL = ObjCreate("ADODB.Connection")
    With $oSQL
      .ConnectionString =("Provider='OraOLEDB.Oracle';" & _
                          "Data Source=" & $sTNS & ";" & _
                          "User Id=" & $sUSERNAME & ";" & _
                          "Password=" & $sPASSWORD & ";")
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
                          "User Id=" & $sUSERNAME & ";" & _
                          "Password=" & $sPASSWORD & ";")
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

    GUICreate("TC X-K?", 120, 120)
    $idChange = GUICtrlCreateButton("Format Pic", 10, 10, 100, 100, $BS_ICON)

    WhoIsThere()
    $sSetFormat = $sFormat
    GUISetState(@SW_SHOW)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop
            Case $idChange
                TheCastling()
                $bSet = False
        EndSwitch
        If $bSet = False Then
            If $sFormat = 'SUPER' Then
                GUICtrlSetImage($idChange, @ScriptDir & "\super.ico")
                $bSet = True
            Else
                GUICtrlSetImage($idChange, @ScriptDir & "\giper.ico")
                $bSet = True
            EndIf
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