' =============================================================================
' frmChat.frm
' Что внутри: UI-форма основного чата: отправка запросов, запуск команд Excel, работа с вложением.
' Количество процедур: 12
' Стиль комментариев: простой язык для быстрой поддержки.
' =============================================================================

Option Explicit

Private chatHistory As String
Private attachedImagePath As String

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub UserForm_Initialize()
    attachedImagePath = ""
    lblAttachment.Caption = ""
    cmbModel.Clear
    cmbModel.AddItem "Gemini 3 Flash"
    cmbModel.AddItem "GPT-5.2"
    cmbModel.AddItem "Gemini 3 Pro"
    cmbModel.AddItem "Claude Sonnet 4.5"
    cmbModel.AddItem "DeepSeek"

    If HasApiKey("claude") Then
        cmbModel.value = "Gemini 3 Flash"
    ElseIf HasApiKey("deepseek") Then
        cmbModel.value = "DeepSeek"
    Else
        cmbModel.ListIndex = 0
    End If

    chkIncludeData.value = True
    chatHistory = "AI: Привет! Я помогу с анализом данных, формулами и форматированием." & vbCrLf & "Выделите данные и опишите задачу." & vbCrLf & "По вопросам доработки обращайтесь: t.me/koladen" & vbCrLf & vbCrLf
    txtChat.value = chatHistory
    lblStatus.Caption = "Готово"

    ' Устанавливаем режим модели
    If IsLocalModelEnabled() And HasLocalModel() Then
        optLocal.value = True
    Else
        optCloud.value = True
    End If
    UpdateModelMode
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub optCloud_Click()
    UpdateModelMode
    SaveLMStudioSetting "Enabled", "0"
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub optLocal_Click()
    UpdateModelMode
    SaveLMStudioSetting "Enabled", "1"
End Sub

' ---
' Что делает: Обновляет состояние интерфейса или данных.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub UpdateModelMode()
    If optLocal.value Then
        cmbModel.Visible = False
        lblLocalModel.Visible = True
        Dim modelName As String
        modelName = GetLMStudioSetting("Model")
        If Len(modelName) = 0 Then
            lblLocalModel.Caption = "LM Studio (авто)"
        ElseIf Len(modelName) > 25 Then
            lblLocalModel.Caption = Left(modelName, 22) & "..."
        Else
            lblLocalModel.Caption = modelName
        End If
    Else
        cmbModel.Visible = True
        lblLocalModel.Visible = False
    End If
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub btnSend_Click()
    Dim userMessage As String
    Dim aiResponse As String
    Dim context As String
    Dim model As String
    Dim useLocal As Boolean

    userMessage = Trim(txtInput.value)
    If Len(userMessage) = 0 Then Exit Sub

    useLocal = optLocal.value

    If useLocal Then
        If Not HasLocalModel() Then
            MsgBox "Настройки LM Studio не заданы. Откройте Настройки.", vbExclamation
            Exit Sub
        End If
    Else
        Select Case cmbModel.value
            Case "DeepSeek"
                model = "deepseek"
            Case "Claude Sonnet 4.5"
                model = "claude"
            Case "GPT-5.2"
                model = "gpt"
            Case "Gemini 3 Pro"
                model = "gemini"
            Case "Gemini 3 Flash"
                model = "gemini-flash"
            Case Else
                model = "deepseek"
        End Select

        Dim keyType As String
        If model = "deepseek" Then
            keyType = "deepseek"
        Else
            keyType = "claude"
        End If

        If Not HasApiKey(keyType) Then
            MsgBox "API-ключ не настроен. Откройте Настройки.", vbExclamation
            Exit Sub
        End If
    End If

    chatHistory = chatHistory & "Вы: " & userMessage & vbCrLf & vbCrLf
    txtChat.value = chatHistory
    txtInput.value = ""

    context = GetWorkbookContext()
    If chkIncludeData.value Then
        context = context & vbCrLf & GetSelectedData()
    End If

    If useLocal Then
        lblStatus.Caption = "LM Studio..."
    Else
        lblStatus.Caption = "Отправка..."
    End If
    Me.Repaint

    If useLocal Then
        aiResponse = SendToLocalAI(userMessage, context)
    Else
        aiResponse = SendToAI(userMessage, model, context, attachedImagePath)
    End If

    attachedImagePath = ""
    lblAttachment.Caption = ""

    Dim commands As String
    Dim execResult As String
    commands = ExtractCommands(aiResponse)

    If Len(commands) > 0 Then
        lblStatus.Caption = "Выполнение..."
        Me.Repaint
        execResult = ExecuteCommands(commands)
        aiResponse = aiResponse & vbCrLf & vbCrLf & "[" & execResult & "]"
    End If

    chatHistory = chatHistory & "AI: " & aiResponse & vbCrLf & vbCrLf
    txtChat.value = chatHistory
    txtChat.SelStart = Len(txtChat.value)
    lblStatus.Caption = "Готово"
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub btnClear_Click()
    chatHistory = "AI: Чат очищен." & vbCrLf & vbCrLf
    txtChat.value = chatHistory
    attachedImagePath = ""
    lblAttachment.Caption = ""
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub btnAttach_Click()
    Dim fd As Object
    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Выберите изображение"
        .Filters.Clear
        .Filters.Add "Изображения", "*.png;*.jpg;*.jpeg;*.gif;*.webp"
        .AllowMultiSelect = False

        If .Show = -1 Then
            attachedImagePath = .SelectedItems(1)
            lblAttachment.Caption = "[IMG] " & GetFileName(attachedImagePath)
        End If
    End With
    Set fd = Nothing
End Sub

' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: Cancel
' Выход: нет (процедура)
' ---
Private Sub lblAttachment_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(attachedImagePath) > 0 Then
        attachedImagePath = ""
        lblAttachment.Caption = ""
    End If
End Sub

' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: fullPath
' Выход: значение функции
' ---
Private Function GetFileName(fullPath As String) As String
    Dim parts() As String
    parts = Split(fullPath, "\")
    GetFileName = parts(UBound(parts))
End Function

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub btnSettings_Click()
    frmSettings.Show vbModal
    UpdateModelMode
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub btnClose_Click()
    Unload Me
End Sub

' ---
' Что делает: Обработчик события формы.
' Вход: KeyCode, Shift
' Выход: нет (процедура)
' ---
Private Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        btnSend_Click
        KeyCode = 0
    End If
End Sub
