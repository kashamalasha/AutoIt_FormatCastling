#AutoIt3Wrapper_Icon=FC.ico                                                     ; заменить иконку для сборки в EXE

#include <ButtonConstants.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>

Global $sFormat                                                                 ; переменная типа String для хранения значения 'FORMAT'

If $CmdLine[0] < 3 Then                                                         ; вывести ошибку, если не введены параметры CMD
   MsgBox(16, @ScriptName & ": Ошибка", _
            "Неверно указана строка соединения с БД" & @CRLF & _
            "Укажи параметры следующим образом:" & @CRLF & _
            @CRLF & _
            """FormatCastling.exe TNS USER PASSWORD""")
   Exit
EndIf

Global $sTNS = $CmdLine[1]                                                      ; переменные параметров командной строки $CmdLine
Global $sUsername = $CmdLine[2]
Global $sPassword = $CmdLine[3]

If @Compiled Then                                                               ; для EXE:
    FileInstall("giper.ico", @TempDir & "\giper.ico")                           ; - скопировать в EXE файлы ресурсов
    FileInstall("super.ico", @TempDir & "\super.ico")                           ;   с распаковкой их во временную папку при выполнении
    FileChangeDir(@TempDir)                                                     ; - сменить рабочую директорию на временную папку
Else                                                                            ; для AU3:
    FileChangeDir(@ScriptDir)                                                   ; - использовать в качестве рабочей директорию папку скрипта
EndIf

; ----------------------------------------------------------------------------
; Функция - WhoIsThere() (Кто там)
;
; Описание:
; выполняет SELECT запрос к серверу для установки текущего значения 'FORMAT'
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
; Функция - TheCastling() (Рокировка)
;
; Описание:
; выполняет UPDATE запрос к серверу для смены текущего значения 'FORMAT'
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
; Функция - Main() (Основная)
; ----------------------------------------------------------------------------
Func Main()
    Local $bSet = False                                                         ; флаг смены состояния для опроса
    Local $lCount = TimerInit()                                                 ; запустить таймер
    Local $sPrevFormat                                                          ; предыдущее значение 'FORMAT'

    GUICreate("FC", 120, 120, @DesktopWidth - 160, 100, Default, $WS_EX_TOPMOST)
    GUISetIcon(@ScriptDir & "\FC.ico")
    $idChange = GUICtrlCreateButton("Format", 10, 10, 100, 100, $BS_ICON)

    WhoIsThere()
    $sPrevFormat = $sFormat
    GUISetState(@SW_SHOW)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                If @Compiled Then                                               ; для EXE - прибраться за собой
                    FileDelete(@TempDir & "/giper.ico")
                    FileDelete(@TempDir & "/super.ico")
                EndIf
                ExitLoop
            Case $idChange
                TheCastling()
                $bSet = False                                                   ; снять флаг bSet, что инициирует смену картинки
        EndSwitch
        If $bSet = False Then                                                   ; смена картинки на кнопке, если флаг bSet снят
            If $sFormat = 'GIPER' Or $sFormat = 'GIPER_MAXI' Then
                GUICtrlSetImage($idChange, @WorkingDir & "\giper.ico")
            Else
                GUICtrlSetImage($idChange, @WorkingDir & "\super.ico")
            EndIf
            $bSet = True
            $sPrevFormat = $sFormat
            GUICtrlSetData($idChange, $sFormat)                                 ; смена названия кнопки (если картинка недоступна)
        EndIf
        If TimerDiff($lCount) >= 1000 Then                                      ; проверка с частотой 1 раз в 1 сек. была ли смена формата
            WhoIsThere()
            If StringCompare($sFormat, $sPrevFormat) <> 0 Then                  ; если кем-то произведена смена формата
                $bSet = False                                                   ; снять флаг bSet, что инициирует смену картинки на следующей итерации
            EndIf
                $lCount = TimerInit()                                           ; перезапустить таймер
        EndIf
    WEnd
EndFunc   ;==>Main

 ;Выполнить сценарий
 Main()