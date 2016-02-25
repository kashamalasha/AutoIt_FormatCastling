#AutoIt3Wrapper_Icon=FC.ico                                                     ; �������� ������ ��� ������ � EXE

#include <ButtonConstants.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>

Global $sFormat                                                                 ; ���������� ���� String ��� �������� �������� 'FORMAT'

If $CmdLine[0] < 3 Then                                                         ; ������� ������, ���� �� ������� ��������� CMD
   MsgBox(16, @ScriptName & ": ������", _
            "������� ������� ������ ���������� � ��" & @CRLF & _
            "����� ��������� ��������� �������:" & @CRLF & _
            @CRLF & _
            """FormatCastling.exe TNS USER PASSWORD""")
   Exit
EndIf

Global $sTNS = $CmdLine[1]                                                      ; ���������� ���������� ��������� ������ $CmdLine
Global $sUsername = $CmdLine[2]
Global $sPassword = $CmdLine[3]

If @Compiled Then                                                               ; ��� EXE:
    FileInstall("giper.ico", @TempDir & "\giper.ico")                           ; - ����������� � EXE ����� ��������
    FileInstall("super.ico", @TempDir & "\super.ico")                           ;   � ����������� �� �� ��������� ����� ��� ����������
    FileChangeDir(@TempDir)                                                     ; - ������� ������� ���������� �� ��������� �����
Else                                                                            ; ��� AU3:
    FileChangeDir(@ScriptDir)                                                   ; - ������������ � �������� ������� ���������� ����� �������
EndIf

; ----------------------------------------------------------------------------
; ������� - WhoIsThere() (��� ���)
;
; ��������:
; ��������� SELECT ������ � ������� ��� ��������� �������� �������� 'FORMAT'
; ----------------------------------------------------------------------------
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

; ----------------------------------------------------------------------------
; ������� - TheCastling() (���������)
;
; ��������:
; ��������� UPDATE ������ � ������� ��� ����� �������� �������� 'FORMAT'
; ----------------------------------------------------------------------------
Func TheCastling()
    Local $oSQL = ObjCreate("ADODB.Connection")

    With $oSQL
      .ConnectionString =("Provider='OraOLEDB.Oracle';" & _
                          "Data Source=" & $sTNS & ";" & _
                          "User Id=" & $sUsername & ";" & _
                          "Password=" & $sPassword & ";")
      .Open
    EndWith

    If $sFormat = 'GIPER' Or $sFormat = 'GIPER_MAXI' Then
        $oSQL.Execute("UPDATE SDD.DEPARTMENT_EXT DE " & _
                      "SET DE.EXT_STRING = 'SUPER' " & _
                      "WHERE DE.EXT_NAME = 'FORMAT' AND " & _
                      "DE.ID_DEPARTMENT = ( " & _
                      "   SELECT D.ID_DEPARTMENT " & _
                      "   FROM SDD.DEPARTMENT D " & _
                      "   WHERE D.IS_HOST = 1)" _
               )
    Else
        $oSQL.Execute("UPDATE SDD.DEPARTMENT_EXT DE " & _
                      "SET DE.EXT_STRING = 'GIPER' " & _
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

; ----------------------------------------------------------------------------
; ������� - Main() (��������)
; ----------------------------------------------------------------------------
Func Main()
    Local $bSet = False                                                         ; ���� ����� ��������� ��� ������
    Local $lCount = TimerInit()                                                 ; ��������� ������
    Local $sPrevFormat                                                          ; ���������� �������� 'FORMAT'

    GUICreate("FC", 120, 120, @DesktopWidth - 160, 100, Default, $WS_EX_TOPMOST)
    GUISetIcon(@ScriptDir & "\FC.ico")
    $idChange = GUICtrlCreateButton("Format", 10, 10, 100, 100, $BS_ICON)

    WhoIsThere()
    $sPrevFormat = $sFormat
    GUISetState(@SW_SHOW)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                If @Compiled Then                                               ; ��� EXE - ���������� �� �����
                    FileDelete(@TempDir & "/giper.ico")
                    FileDelete(@TempDir & "/super.ico")
                EndIf
                ExitLoop
            Case $idChange
                TheCastling()
                $bSet = False                                                   ; ����� ���� bSet, ��� ���������� ����� ��������
        EndSwitch
        If $bSet = False Then                                                   ; ����� �������� �� ������, ���� ���� bSet ����
            If $sFormat = 'GIPER' Or $sFormat = 'GIPER_MAXI' Then
                GUICtrlSetImage($idChange, @WorkingDir & "\giper.ico")
            Else
                GUICtrlSetImage($idChange, @WorkingDir & "\super.ico")
            EndIf
            $bSet = True
            $sPrevFormat = $sFormat
            GUICtrlSetData($idChange, $sFormat)                                 ; ����� �������� ������ (���� �������� ����������)
        EndIf
        If TimerDiff($lCount) >= 1000 Then                                      ; �������� � �������� 1 ��� � 1 ���. ���� �� ����� �������
            WhoIsThere()
            If StringCompare($sFormat, $sPrevFormat) <> 0 Then                  ; ���� ���-�� ����������� ����� �������
                $bSet = False                                                   ; ����� ���� bSet, ��� ���������� ����� �������� �� ��������� ��������
            EndIf
                $lCount = TimerInit()                                           ; ������������� ������
        EndIf
    WEnd
EndFunc   ;==>Main

 ;��������� ��������
 Main()