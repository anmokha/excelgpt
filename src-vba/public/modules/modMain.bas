' =============================================================================
' modMain.bas
' Что внутри: Точки входа add-in: кнопка в меню Excel, запуск форм, быстрый сценарий анализа.
' Количество процедур: 7
' Стиль комментариев: простой язык для быстрой поддержки.
' =============================================================================

'========================================
' Главный модуль - точки входа
'========================================

Option Explicit

'----------------------------------------
' Показать окно чата
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: нет
' Выход: нет (процедура)
' ---
Public Sub ShowAIAssistant()
    frmChat.Show vbModeless
End Sub

'----------------------------------------
' Показать настройки
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: нет
' Выход: нет (процедура)
' ---
Public Sub ShowSettings()
    frmSettings.Show vbModal
End Sub

'----------------------------------------
' Создание меню при загрузке надстройки
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: нет
' Выход: нет (процедура)
' ---
Public Sub Auto_Open()
    CreateMenu
End Sub

'----------------------------------------
' Удаление меню при выгрузке
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: нет
' Выход: нет (процедура)
' ---
Public Sub Auto_Close()
    DeleteMenu
End Sub

'----------------------------------------
' Создание пункта меню
'----------------------------------------
' ---
' Что делает: Создаёт новый объект или структуру в Excel.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub CreateMenu()
    On Error Resume Next

    Dim cmdBar As CommandBar
    Dim cmdBtn As CommandBarButton

    ' Удаляем старое меню если есть
    DeleteMenu

    ' Для Excel 2007+ добавляем на вкладку "Надстройки"
    Set cmdBar = Application.CommandBars("Worksheet Menu Bar")

    ' Добавляем кнопку
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton, Temporary:=True)

    With cmdBtn
        .Caption = "AI Ассистент"
        .Style = msoButtonCaption
        .OnAction = "ShowAIAssistant"
        .Tag = "AIAssistantButton"
    End With
End Sub

'----------------------------------------
' Удаление меню
'----------------------------------------
' ---
' Что делает: Удаляет объект или очищает данные.
' Вход: нет
' Выход: нет (процедура)
' ---
Private Sub DeleteMenu()
    On Error Resume Next

    Dim ctrl As CommandBarControl

    For Each ctrl In Application.CommandBars("Worksheet Menu Bar").Controls
        If ctrl.Tag = "AIAssistantButton" Then
            ctrl.Delete
        End If
    Next ctrl
End Sub

'----------------------------------------
' Быстрый анализ выделенных данных
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: нет
' Выход: нет (процедура)
' ---
Public Sub QuickAnalyze()
    Dim selectedData As String
    Dim context As String
    Dim response As String
    Dim model As String

    selectedData = GetSelectedData()

    If Len(selectedData) = 0 Then
        MsgBox "Сначала выделите данные для анализа", vbExclamation, "AI Ассистент"
        Exit Sub
    End If

    context = GetWorkbookContext() & vbCrLf & selectedData

    ' Используем DeepSeek по умолчанию
    model = "deepseek"
    If Not HasApiKey(model) Then
        model = "claude"
    End If

    If Not HasApiKey(model) Then
        MsgBox "API-ключи не настроены. Откройте настройки.", vbExclamation, "AI Ассистент"
        ShowSettings
        Exit Sub
    End If

    Application.StatusBar = "AI анализирует данные..."

    response = SendToAI("Проанализируй эти данные и найди проблемы с форматированием или качеством данных:", model, context)

    Application.StatusBar = False

    MsgBox response, vbInformation, "Результат анализа"
End Sub
