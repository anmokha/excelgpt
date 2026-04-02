' =============================================================================
' modExcelHelper.bas
' Что внутри: Работа с Excel-контекстом: анализ книги, выборки данных и вспомогательные Excel-утилиты.
' Количество процедур: 13
' Стиль комментариев: простой язык для быстрой поддержки.
' =============================================================================

'========================================
' Модуль для работы с Excel
'========================================

Option Explicit

'----------------------------------------
' Получение названия версии Excel
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: нет
' Выход: значение функции
' ---
Private Function GetExcelVersionName() As String
    Dim ver As String
    ver = Application.Version

    Select Case ver
        Case "12.0"
            GetExcelVersionName = "Excel 2007 - используй только функции, доступные в Excel 2007"
        Case "14.0"
            GetExcelVersionName = "Excel 2010 - используй только функции, доступные в Excel 2010"
        Case "15.0"
            GetExcelVersionName = "Excel 2013 - используй только функции, доступные в Excel 2013"
        Case "16.0"
            GetExcelVersionName = "Excel 2016/2019/365"
        Case Else
            GetExcelVersionName = "Excel " & ver
    End Select
End Function

'----------------------------------------
' Получение контекста текущей книги
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: нет
' Выход: значение функции
' ---
Public Function GetWorkbookContext() As String
    On Error Resume Next

    Dim context As String
    Dim ws As Worksheet
    Dim usedRng As Range
    Dim headers As String
    Dim sampleData As String
    Dim i As Long, j As Long
    Dim firstRow As Long
    Dim firstCol As Long
    Dim lastRow As Long
    Dim lastCol As Long

    ' Добавляем версию Excel
    context = "Версия Excel: " & GetExcelVersionName() & vbCrLf
    context = context & "Активный лист: " & ActiveSheet.Name & vbCrLf
    context = context & "Листы в книге: "

    For Each ws In ActiveWorkbook.Worksheets
        context = context & ws.Name & ", "
    Next ws
    context = Left(context, Len(context) - 2) & vbCrLf

    ' Используемый диапазон
    Set usedRng = ActiveSheet.UsedRange
    If Not usedRng Is Nothing Then
        ' Получаем точные координаты
        firstRow = usedRng.Row
        firstCol = usedRng.Column
        lastRow = firstRow + usedRng.Rows.Count - 1
        lastCol = firstCol + usedRng.columns.Count - 1

        context = context & "Используемый диапазон: " & usedRng.address & vbCrLf
        context = context & "Начинается со строки " & firstRow & ", столбца " & ColLetter(firstCol) & vbCrLf
        context = context & "Заканчивается строкой " & lastRow & ", столбцом " & ColLetter(lastCol) & vbCrLf
        context = context & "Всего строк: " & usedRng.Rows.Count & ", столбцов: " & usedRng.columns.Count & vbCrLf

        ' Заголовки (первая строка данных с их адресами)
        If usedRng.Rows.Count > 0 Then
            context = context & vbCrLf & "Структура данных (первая строка - заголовки):" & vbCrLf
            For j = 1 To Application.Min(usedRng.columns.Count, 10)
                Dim cellAddr As String
                cellAddr = ColLetter(firstCol + j - 1) & firstRow
                context = context & "  " & cellAddr & ": " & usedRng.Cells(1, j).value & vbCrLf
            Next j

            ' Добавляем пример данных (вторая строка)
            If usedRng.Rows.Count > 1 Then
                context = context & vbCrLf & "Пример данных (строка " & (firstRow + 1) & "):" & vbCrLf
                For j = 1 To Application.Min(usedRng.columns.Count, 10)
                    cellAddr = ColLetter(firstCol + j - 1) & (firstRow + 1)
                    context = context & "  " & cellAddr & ": " & usedRng.Cells(2, j).value & vbCrLf
                Next j
            End If
        End If
    End If

    GetWorkbookContext = context
End Function

'----------------------------------------
' Преобразование номера столбца в букву
'----------------------------------------
' ---
' Что делает: Вспомогательная функция модуля.
' Вход: colNum
' Выход: значение функции
' ---
Private Function ColLetter(colNum As Long) As String
    Dim result As String
    Dim n As Long

    n = colNum
    result = ""

    Do While n > 0
        result = Chr(((n - 1) Mod 26) + 65) & result
        n = (n - 1) \ 26
    Loop

    ColLetter = result
End Function

'----------------------------------------
' Получение выделенных данных
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: нет
' Выход: значение функции
' ---
Public Function GetSelectedData() As String
    On Error Resume Next

    Dim sel As Range
    Dim result As String
    Dim i As Long, j As Long
    Dim maxRows As Long
    Dim maxCols As Long
    Dim firstRow As Long
    Dim firstCol As Long
    Dim lastRow As Long
    Dim lastCol As Long

    Set sel = Selection

    If sel Is Nothing Then
        GetSelectedData = ""
        Exit Function
    End If

    If TypeName(sel) <> "Range" Then
        GetSelectedData = ""
        Exit Function
    End If

    ' Получаем точные координаты
    firstRow = sel.Row
    firstCol = sel.Column
    lastRow = firstRow + sel.Rows.Count - 1
    lastCol = firstCol + sel.columns.Count - 1

    result = "=== ВЫДЕЛЕННЫЕ ДАННЫЕ ===" & vbCrLf
    result = result & "Диапазон: " & sel.address & vbCrLf
    result = result & "Первая ячейка: " & ColLetter(firstCol) & firstRow & vbCrLf
    result = result & "Последняя ячейка: " & ColLetter(lastCol) & lastRow & vbCrLf
    result = result & "Размер: " & sel.Rows.Count & " строк x " & sel.columns.Count & " столбцов" & vbCrLf & vbCrLf

    ' Ограничиваем количество данных
    maxRows = Application.Min(sel.Rows.Count, 30)
    maxCols = Application.Min(sel.columns.Count, 10)

    ' Выводим данные с адресами строк
    result = result & "Данные (с адресами):" & vbCrLf

    ' Заголовок таблицы с буквами столбцов
    result = result & "Строка" & vbTab
    For j = 1 To maxCols
        result = result & ColLetter(firstCol + j - 1) & vbTab
    Next j
    result = result & vbCrLf

    ' Данные
    For i = 1 To maxRows
        result = result & (firstRow + i - 1) & vbTab
        For j = 1 To maxCols
            result = result & sel.Cells(i, j).value
            If j < maxCols Then result = result & vbTab
        Next j
        result = result & vbCrLf
    Next i

    If sel.Rows.Count > maxRows Then
        result = result & "... и ещё " & (sel.Rows.Count - maxRows) & " строк" & vbCrLf
    End If

    ' Подсказка для AI
    result = result & vbCrLf & "ВАЖНО: Используй ТОЧНЫЕ адреса ячеек из данных выше!" & vbCrLf
    result = result & "Первая строка данных: " & firstRow & ", последняя: " & lastRow & vbCrLf
    result = result & "Первый столбец: " & ColLetter(firstCol) & ", последний: " & ColLetter(lastCol) & vbCrLf

    GetSelectedData = result
End Function

'----------------------------------------
' Установка значения в ячейку
'----------------------------------------
' ---
' Что делает: Записывает значение в Excel или в настройки.
' Вход: address, value
' Выход: нет (процедура)
' ---
Public Sub SetCellValue(address As String, value As Variant)
    On Error Resume Next
    ActiveSheet.Range(address).value = value
End Sub

'----------------------------------------
' Установка формулы в ячейку
'----------------------------------------
' ---
' Что делает: Записывает значение в Excel или в настройки.
' Вход: address, formula
' Выход: нет (процедура)
' ---
Public Sub SetCellFormula(address As String, formula As String)
    On Error Resume Next
    ActiveSheet.Range(address).formula = formula
End Sub

'----------------------------------------
' Установка значений в диапазон
'----------------------------------------
' ---
' Что делает: Записывает значение в Excel или в настройки.
' Вход: address, values
' Выход: нет (процедура)
' ---
Public Sub SetRangeValues(address As String, values As Variant)
    On Error Resume Next
    ActiveSheet.Range(address).value = values
End Sub

'----------------------------------------
' Форматирование диапазона
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: address, formatType, formatValue
' Выход: нет (процедура)
' ---
Public Sub FormatRange(address As String, formatType As String, formatValue As String)
    On Error Resume Next

    Dim rng As Range
    Set rng = ActiveSheet.Range(address)

    Select Case formatType
        Case "numberformat"
            rng.NumberFormat = formatValue
        Case "bold"
            rng.Font.Bold = (formatValue = "true")
        Case "italic"
            rng.Font.Italic = (formatValue = "true")
        Case "fontcolor"
            rng.Font.Color = CLng(formatValue)
        Case "fillcolor"
            rng.Interior.Color = CLng(formatValue)
        Case "fontsize"
            rng.Font.Size = CInt(formatValue)
        Case "align"
            Select Case formatValue
                Case "left": rng.HorizontalAlignment = xlLeft
                Case "center": rng.HorizontalAlignment = xlCenter
                Case "right": rng.HorizontalAlignment = xlRight
            End Select
    End Select
End Sub

'----------------------------------------
' Автоподбор ширины столбцов
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: address
' Выход: нет (процедура)
' ---
Public Sub AutoFitColumns(Optional address As String = "")
    On Error Resume Next

    If Len(address) > 0 Then
        ActiveSheet.Range(address).columns.AutoFit
    Else
        ActiveSheet.UsedRange.columns.AutoFit
    End If
End Sub

'----------------------------------------
' Сортировка диапазона
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: address, columnIndex, ascending
' Выход: нет (процедура)
' ---
Public Sub SortRange(address As String, columnIndex As Long, ascending As Boolean)
    On Error Resume Next

    Dim rng As Range
    Set rng = ActiveSheet.Range(address)

    rng.Sort Key1:=rng.columns(columnIndex), _
             Order1:=IIf(ascending, xlAscending, xlDescending), _
             Header:=xlYes
End Sub

'----------------------------------------
' Создание таблицы
'----------------------------------------
' ---
' Что делает: Создаёт новый объект или структуру в Excel.
' Вход: address, tableName
' Выход: нет (процедура)
' ---
Public Sub CreateTable(address As String, tableName As String)
    On Error Resume Next

    Dim rng As Range
    Dim tbl As ListObject

    Set rng = ActiveSheet.Range(address)
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)

    If Len(tableName) > 0 Then
        tbl.Name = tableName
    End If

    tbl.TableStyle = "TableStyleMedium2"
End Sub

'----------------------------------------
' Удаление дубликатов
'----------------------------------------
' ---
' Что делает: Удаляет объект или очищает данные.
' Вход: address, columns
' Выход: нет (процедура)
' ---
Public Sub RemoveDuplicates(address As String, Optional columns As String = "")
    On Error Resume Next

    Dim rng As Range
    Set rng = ActiveSheet.Range(address)

    If Len(columns) = 0 Then
        rng.RemoveDuplicates columns:=1, Header:=xlYes
    Else
        ' Парсинг столбцов из строки "1,2,3"
        Dim colArr() As String
        Dim colNums() As Long
        Dim i As Long

        colArr = Split(columns, ",")
        ReDim colNums(UBound(colArr))

        For i = 0 To UBound(colArr)
            colNums(i) = CLng(Trim(colArr(i)))
        Next i

        rng.RemoveDuplicates columns:=colNums, Header:=xlYes
    End If
End Sub

'----------------------------------------
' Поиск и замена
'----------------------------------------
' ---
' Что делает: Вспомогательная процедура модуля.
' Вход: findText, replaceText, address
' Выход: нет (процедура)
' ---
Public Sub FindAndReplace(findText As String, replaceText As String, Optional address As String = "")
    On Error Resume Next

    Dim rng As Range

    If Len(address) > 0 Then
        Set rng = ActiveSheet.Range(address)
    Else
        Set rng = ActiveSheet.UsedRange
    End If

    rng.Replace What:=findText, Replacement:=replaceText, LookAt:=xlPart
End Sub
