' =============================================================================
' frmSettings.frm
' Что внутри: UI-форма настроек: API-ключи, параметры LM Studio, выбор локальной модели.
' Количество процедур: 5
' Стиль комментариев: простой язык для быстрой поддержки.
' =============================================================================

Option Explicit

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub UserForm_Initialize()
    txtDeepSeekKey.value = GetApiKey("DeepSeekKey")
    txtOpenRouterKey.value = GetApiKey("OpenRouterKey")

    txtLMStudioIP.value = GetLMStudioSetting("IP")
    txtLMStudioPort.value = GetLMStudioSetting("Port")

    LoadLMStudioModels

    Dim savedModel As String
    savedModel = GetLMStudioSetting("Model")
    If Len(savedModel) > 0 Then
        On Error Resume Next
        cmbLMStudioModel.value = savedModel
        On Error GoTo 0
    End If
End Sub

' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub LoadLMStudioModels()
    On Error Resume Next
    cmbLMStudioModel.Clear
    cmbLMStudioModel.AddItem "(авто)"

    Dim models As String
    Dim modelArr() As String
    Dim i As Long

    models = GetLMStudioModels()

    If Left(models, 6) <> "ERROR:" And Len(models) > 0 Then
        modelArr = Split(models, "|")
        For i = LBound(modelArr) To UBound(modelArr)
            If Len(Trim(modelArr(i))) > 0 Then
                cmbLMStudioModel.AddItem modelArr(i)
            End If
        Next i
    End If
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub btnRefreshModels_Click()
    lblLMStatus.Caption = "Загрузка..."
    Me.Repaint

    If IsLMStudioAvailable() Then
        LoadLMStudioModels
        lblLMStatus.Caption = "LM Studio OK"
        lblLMStatus.ForeColor = &H8000&
    Else
        lblLMStatus.Caption = "Недоступен"
        lblLMStatus.ForeColor = &HFF&
    End If
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub btnSave_Click()
    SaveApiKey "DeepSeekKey", Trim(txtDeepSeekKey.value)
    SaveApiKey "OpenRouterKey", Trim(txtOpenRouterKey.value)

    SaveLMStudioSetting "IP", Trim(txtLMStudioIP.value)
    SaveLMStudioSetting "Port", Trim(txtLMStudioPort.value)

    Dim selectedModel As String
    selectedModel = cmbLMStudioModel.value
    If InStr(selectedModel, "(авто") > 0 Then selectedModel = ""
    SaveLMStudioSetting "Model", selectedModel

    MsgBox "Настройки сохранены!", vbInformation
    Unload Me
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub btnCancel_Click()
    Unload Me
End Sub
