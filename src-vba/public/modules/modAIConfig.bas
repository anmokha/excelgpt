' =============================================================================
' modAIConfig.bas
' Что внутри: Доступ к API-ключам и настройкам LM Studio через реестр Windows.
' Автоматически собран из модульной структуры проекта.
' =============================================================================

Option Explicit

' Настройки LM Studio по умолчанию
Private Const LMSTUDIO_DEFAULT_IP As String = "127.0.0.1"
Private Const LMSTUDIO_DEFAULT_PORT As String = "1234"

' Хранение ключей (в реестре)
Private Const REG_PATH As String = "HKEY_CURRENT_USER\Software\ExcelAIAssistant\"


'----------------------------------------
' Получение API-ключа из реестра
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: keyName
' Выход: значение функции
' ---
Public Function GetApiKey(keyName As String) As String
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetApiKey = wsh.RegRead(REG_PATH & keyName)
    Set wsh = Nothing
End Function


'----------------------------------------
' Сохранение API-ключа в реестр
'----------------------------------------
' ---
' Что делает: Сохраняет данные в постоянное хранилище.
' Вход: keyName, keyValue
' Выход: нет (процедура)
' ---
Public Sub SaveApiKey(keyName As String, keyValue As String)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite REG_PATH & keyName, keyValue, "REG_SZ"
    Set wsh = Nothing
End Sub


'----------------------------------------
' Проверка наличия ключа
'----------------------------------------
' ---
' Что делает: Проверяет условие и возвращает True/False.
' Вход: model
' Выход: значение функции
' ---
Public Function HasApiKey(model As String) As Boolean
    If model = "deepseek" Then
        HasApiKey = Len(GetApiKey("DeepSeekKey")) > 0
    Else
        HasApiKey = Len(GetApiKey("OpenRouterKey")) > 0
    End If
End Function


'========================================
' ФУНКЦИИ ДЛЯ РАБОТЫ С LM STUDIO
'========================================

'----------------------------------------
' Получение настройки LM Studio из реестра
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: settingName
' Выход: значение функции
' ---
Public Function GetLMStudioSetting(settingName As String) As String
    On Error Resume Next
    Dim wsh As Object
    Dim result As String
    Set wsh = CreateObject("WScript.Shell")
    result = wsh.RegRead(REG_PATH & "LMStudio_" & settingName)
    Set wsh = Nothing

    ' Значения по умолчанию
    If Len(result) = 0 Then
        Select Case settingName
            Case "IP": result = LMSTUDIO_DEFAULT_IP
            Case "Port": result = LMSTUDIO_DEFAULT_PORT
            Case "Model": result = ""
            Case "Enabled": result = "0"
        End Select
    End If

    GetLMStudioSetting = result
End Function


'----------------------------------------
' Сохранение настройки LM Studio в реестр
'----------------------------------------
' ---
' Что делает: Сохраняет данные в постоянное хранилище.
' Вход: settingName, settingValue
' Выход: нет (процедура)
' ---
Public Sub SaveLMStudioSetting(settingName As String, settingValue As String)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite REG_PATH & "LMStudio_" & settingName, settingValue, "REG_SZ"
    Set wsh = Nothing
End Sub


'----------------------------------------
' Проверка включена ли локальная модель
'----------------------------------------
' ---
' Что делает: Проверяет условие и возвращает True/False.
' Вход: нет
' Выход: значение функции
' ---
Public Function IsLocalModelEnabled() As Boolean
    IsLocalModelEnabled = (GetLMStudioSetting("Enabled") = "1")
End Function


'----------------------------------------
' Проверка настроена ли локальная модель
'----------------------------------------
' ---
' Что делает: Проверяет условие и возвращает True/False.
' Вход: нет
' Выход: значение функции
' ---
Public Function HasLocalModel() As Boolean
    Dim ip As String
    Dim port As String
    ip = GetLMStudioSetting("IP")
    port = GetLMStudioSetting("Port")
    HasLocalModel = (Len(ip) > 0 And Len(port) > 0)
End Function
