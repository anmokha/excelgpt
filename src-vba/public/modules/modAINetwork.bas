' =============================================================================
' modAINetwork.bas
' Что внутри: Запросы к облачным моделям и LM Studio, сбор JSON, HTTP и парсинг ответов.
' Автоматически собран из модульной структуры проекта.
' =============================================================================

Option Explicit

' Константы API
Private Const DEEPSEEK_URL As String = "https://api.deepseek.com/chat/completions"
Private Const OPENROUTER_URL As String = "https://openrouter.ai/api/v1/chat/completions"
Private Const DEEPSEEK_MODEL As String = "deepseek-chat"
Private Const CLAUDE_MODEL As String = "anthropic/claude-4.5-sonnet-20250929"
Private Const GPT_MODEL As String = "openai/gpt-5.2"
Private Const GEMINI_MODEL As String = "google/gemini-3-pro-preview"
Private Const GEMINI_FLASH_MODEL As String = "google/gemini-3-flash-preview-20251217"


'----------------------------------------
' Отправка запроса к AI
'----------------------------------------
' ---
' Что делает: Отправляет запрос во внешний сервис и получает ответ.
' Вход: userMessage, model, excelContext, imagePath
' Выход: значение функции
' ---
Public Function SendToAI(userMessage As String, model As String, Optional excelContext As String = "", Optional imagePath As String = "") As String
    On Error GoTo ErrorHandler

    Dim apiUrl As String
    Dim apiKey As String
    Dim modelName As String
    Dim requestBody As String
    Dim response As String
    Dim imageBase64 As String

    ' Проверяем изображение
    imageBase64 = ""
    If Len(imagePath) > 0 Then
        ' DeepSeek не поддерживает изображения
        If model = "deepseek" Then
            SendToAI = "ОШИБКА: DeepSeek не поддерживает изображения. Выберите другую модель (Claude, GPT, Gemini)."
            Exit Function
        End If
        imageBase64 = ImageToBase64(imagePath)
        If Left(imageBase64, 6) = "ERROR:" Then
            SendToAI = "ОШИБКА загрузки изображения: " & Mid(imageBase64, 7)
            Exit Function
        End If
    End If

    ' Выбор API
    If model = "deepseek" Then
        apiUrl = DEEPSEEK_URL
        apiKey = GetApiKey("DeepSeekKey")
        modelName = DEEPSEEK_MODEL
    Else
        ' Все остальные модели через OpenRouter
        apiUrl = OPENROUTER_URL
        apiKey = GetApiKey("OpenRouterKey")

        Select Case model
            Case "claude"
                modelName = CLAUDE_MODEL
            Case "gpt"
                modelName = GPT_MODEL
            Case "gemini"
                modelName = GEMINI_MODEL
            Case "gemini-flash"
                modelName = GEMINI_FLASH_MODEL
            Case Else
                modelName = CLAUDE_MODEL
        End Select
    End If

    If Len(apiKey) = 0 Then
        SendToAI = "ОШИБКА: API-ключ не настроен. Откройте настройки."
        Exit Function
    End If

    ' Формируем системный промпт
    Dim systemPrompt As String
    systemPrompt = BuildSystemPrompt(excelContext)

    ' Формируем JSON запрос (с изображением или без)
    If Len(imageBase64) > 0 Then
        requestBody = BuildRequestJSONWithImage(systemPrompt, userMessage, modelName, imageBase64, imagePath)
    Else
        requestBody = BuildRequestJSON(systemPrompt, userMessage, modelName)
    End If

    ' Отправляем запрос
    response = SendHTTPRequest(apiUrl, apiKey, requestBody, model)

    ' Парсим ответ
    SendToAI = ParseResponse(response)
    Exit Function

ErrorHandler:
    SendToAI = "ОШИБКА: " & Err.Description
End Function


'----------------------------------------
' Конвертация изображения в Base64
'----------------------------------------
' ---
' Что делает: Вспомогательная функция модуля.
' Вход: filePath
' Выход: значение функции
' ---
Private Function ImageToBase64(filePath As String) As String
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    Dim fileData() As Byte
    Dim fileLen As Long

    ' Проверяем существование файла
    If Dir(filePath) = "" Then
        ImageToBase64 = "ERROR:Файл не найден"
        Exit Function
    End If

    ' Читаем файл
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    fileLen = LOF(fileNum)

    If fileLen = 0 Then
        Close #fileNum
        ImageToBase64 = "ERROR:Файл пуст"
        Exit Function
    End If

    If fileLen > 20000000 Then ' 20MB лимит
        Close #fileNum
        ImageToBase64 = "ERROR:Файл слишком большой (макс 20MB)"
        Exit Function
    End If

    ReDim fileData(fileLen - 1)
    Get #fileNum, , fileData
    Close #fileNum

    ' Конвертируем в Base64
    ImageToBase64 = EncodeBase64(fileData)
    Exit Function

ErrorHandler:
    ImageToBase64 = "ERROR:" & Err.Description
End Function


'----------------------------------------
' Кодирование массива байт в Base64
'----------------------------------------
' ---
' Что делает: Вспомогательная функция модуля.
' Вход: arrData()
' Выход: значение функции
' ---
Private Function EncodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
End Function


'----------------------------------------
' Получение MIME-типа по расширению файла
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: filePath
' Выход: значение функции
' ---
Private Function GetMimeType(filePath As String) As String
    Dim ext As String
    ext = LCase(Right(filePath, 4))

    Select Case ext
        Case ".png"
            GetMimeType = "image/png"
        Case ".jpg", "jpeg"
            GetMimeType = "image/jpeg"
        Case ".gif"
            GetMimeType = "image/gif"
        Case "webp"
            GetMimeType = "image/webp"
        Case Else
            GetMimeType = "image/png"
    End Select
End Function


'----------------------------------------
' Формирование JSON с изображением (Vision API)
'----------------------------------------
' ---
' Что делает: Собирает служебную строку или JSON для следующего шага.
' Вход: systemPrompt, userMessage, modelName, imageBase64, imagePath
' Выход: значение функции
' ---
Private Function BuildRequestJSONWithImage(systemPrompt As String, userMessage As String, modelName As String, imageBase64 As String, imagePath As String) As String
    Dim json As String
    Dim mimeType As String

    mimeType = GetMimeType(imagePath)

    json = "{" & vbCrLf
    json = json & "  ""model"": """ & modelName & """," & vbCrLf
    json = json & "  ""messages"": [" & vbCrLf
    json = json & "    {""role"": ""system"", ""content"": """ & EscapeJSON(systemPrompt) & """}," & vbCrLf
    json = json & "    {""role"": ""user"", ""content"": [" & vbCrLf
    json = json & "      {""type"": ""image_url"", ""image_url"": {""url"": ""data:" & mimeType & ";base64," & imageBase64 & """}}," & vbCrLf
    json = json & "      {""type"": ""text"", ""text"": """ & EscapeJSON(userMessage) & """}" & vbCrLf
    json = json & "    ]}" & vbCrLf
    json = json & "  ]," & vbCrLf
    json = json & "  ""max_tokens"": 4096" & vbCrLf
    json = json & "}"

    BuildRequestJSONWithImage = json
End Function


'----------------------------------------
' Построение системного промпта
'----------------------------------------
' ---
' Что делает: Собирает служебную строку или JSON для следующего шага.
' Вход: excelContext
' Выход: значение функции
' ---
Private Function BuildSystemPrompt(excelContext As String) As String
    Dim prompt As String

    prompt = "Ты — AI-ассистент для Microsoft Excel. Выполняешь действия с данными АВТОМАТИЧЕСКИ." & vbCrLf & vbCrLf
    prompt = prompt & "КРИТИЧЕСКИ ВАЖНО:" & vbCrLf
    prompt = prompt & "1. Возвращай команды для автоматического выполнения!" & vbCrLf
    prompt = prompt & "2. НЕ давай инструкции пользователю — ВЫПОЛНЯЙ сам!" & vbCrLf
    prompt = prompt & "3. ВСЕГДА проверяй адреса из контекста! Заголовки могут быть НЕ в строке 1!" & vbCrLf
    prompt = prompt & "4. Используй ТОЧНЫЕ адреса из раздела 'Контекст Excel' ниже!" & vbCrLf
    prompt = prompt & "5. ФОРМУЛЫ: пиши на АНГЛИЙСКОМ языке (SUM, IF, VLOOKUP...), разделитель запятая. Система автоматически переведёт." & vbCrLf & vbCrLf
    prompt = prompt & "Формат ответа:" & vbCrLf
    prompt = prompt & "1. Краткое описание (1-2 предложения)" & vbCrLf
    prompt = prompt & "2. ОБЯЗАТЕЛЬНО блок команд:" & vbCrLf & vbCrLf
    prompt = prompt & "```commands" & vbCrLf
    prompt = prompt & "SET_VALUE|<адрес>|значение" & vbCrLf
    prompt = prompt & "SET_FORMULA|<адрес>|формула" & vbCrLf
    prompt = prompt & "```" & vbCrLf & vbCrLf
    prompt = prompt & "ПРАВИЛО АДРЕСОВ: Если данные начинаются со строки N, то:" & vbCrLf
    prompt = prompt & "- Новый заголовок добавляй в строку N (рядом с существующими заголовками)" & vbCrLf
    prompt = prompt & "- Формулы начинай со строки N+1 (где первая строка данных)" & vbCrLf
    prompt = prompt & "Пример: если заголовки в строке 2, данные в строках 3-5, то:" & vbCrLf
    prompt = prompt & "  SET_VALUE|E2|НовыйЗаголовок (в строку 2, к заголовкам)" & vbCrLf
    prompt = prompt & "  SET_FORMULA|E3|=B3*C3 (в строку 3, к данным)" & vbCrLf
    prompt = prompt & "  FILL_DOWN|E3|E5 (протянуть до последней строки данных)" & vbCrLf & vbCrLf

    ' === ПОЛНЫЙ СПИСОК КОМАНД ===
    prompt = prompt & "=== ДОСТУПНЫЕ КОМАНДЫ ===" & vbCrLf & vbCrLf

    ' РАБОТА С ЯЧЕЙКАМИ
    prompt = prompt & "--- ЯЧЕЙКИ ---" & vbCrLf
    prompt = prompt & "SET_VALUE|адрес|значение - записать значение" & vbCrLf
    prompt = prompt & "SET_FORMULA|адрес|формула - записать формулу" & vbCrLf
    prompt = prompt & "FILL_DOWN|начало|конец - протянуть вниз" & vbCrLf
    prompt = prompt & "FILL_RIGHT|начало|конец - протянуть вправо" & vbCrLf
    prompt = prompt & "FILL_SERIES|диапазон|шаг - заполнить последовательность" & vbCrLf
    prompt = prompt & "CLEAR_CONTENTS|диапазон - очистить содержимое" & vbCrLf
    prompt = prompt & "CLEAR_FORMAT|диапазон - очистить форматирование" & vbCrLf
    prompt = prompt & "CLEAR_ALL|диапазон - очистить всё" & vbCrLf
    prompt = prompt & "COPY|откуда|куда - копировать" & vbCrLf
    prompt = prompt & "CUT|откуда|куда - вырезать" & vbCrLf
    prompt = prompt & "PASTE_VALUES|откуда|куда - вставить значения" & vbCrLf
    prompt = prompt & "TRANSPOSE|откуда|куда - транспонировать" & vbCrLf & vbCrLf

    ' ФОРМАТИРОВАНИЕ
    prompt = prompt & "--- ФОРМАТИРОВАНИЕ ---" & vbCrLf
    prompt = prompt & "BOLD|диапазон - жирный" & vbCrLf
    prompt = prompt & "ITALIC|диапазон - курсив" & vbCrLf
    prompt = prompt & "UNDERLINE|диапазон - подчёркивание" & vbCrLf
    prompt = prompt & "STRIKETHROUGH|диапазон - зачёркивание" & vbCrLf
    prompt = prompt & "FONT_NAME|диапазон|имя_шрифта - шрифт" & vbCrLf
    prompt = prompt & "FONT_SIZE|диапазон|размер - размер шрифта" & vbCrLf
    prompt = prompt & "FONT_COLOR|диапазон|цвет - цвет текста (RED,GREEN,BLUE,BLACK,WHITE,YELLOW,ORANGE,PURPLE,GRAY или RGB:255,0,0)" & vbCrLf
    prompt = prompt & "FILL_COLOR|диапазон|цвет - заливка" & vbCrLf
    prompt = prompt & "BORDER|диапазон|стиль - границы (ALL,TOP,BOTTOM,LEFT,RIGHT,NONE)" & vbCrLf
    prompt = prompt & "BORDER_THICK|диапазон - толстые границы" & vbCrLf
    prompt = prompt & "ALIGN_H|диапазон|выравнивание - горизонтальное (LEFT,CENTER,RIGHT)" & vbCrLf
    prompt = prompt & "ALIGN_V|диапазон|выравнивание - вертикальное (TOP,CENTER,BOTTOM)" & vbCrLf
    prompt = prompt & "WRAP_TEXT|диапазон - перенос текста" & vbCrLf
    prompt = prompt & "MERGE|диапазон - объединить ячейки" & vbCrLf
    prompt = prompt & "UNMERGE|диапазон - разъединить ячейки" & vbCrLf
    prompt = prompt & "FORMAT_NUMBER|диапазон|формат - формат числа (#,##0.00)" & vbCrLf
    prompt = prompt & "FORMAT_DATE|диапазон|формат - формат даты (DD.MM.YYYY)" & vbCrLf
    prompt = prompt & "FORMAT_PERCENT|диапазон - формат процентов" & vbCrLf
    prompt = prompt & "FORMAT_CURRENCY|диапазон|символ - формат валюты" & vbCrLf
    prompt = prompt & "AUTOFIT|диапазон - автоширина столбцов" & vbCrLf
    prompt = prompt & "AUTOFIT_ROWS|диапазон - автовысота строк" & vbCrLf
    prompt = prompt & "COLUMN_WIDTH|столбец|ширина - ширина столбца" & vbCrLf
    prompt = prompt & "ROW_HEIGHT|строка|высота - высота строки" & vbCrLf & vbCrLf

    ' СТРОКИ И СТОЛБЦЫ
    prompt = prompt & "--- СТРОКИ И СТОЛБЦЫ ---" & vbCrLf
    prompt = prompt & "INSERT_ROW|номер - вставить строку" & vbCrLf
    prompt = prompt & "INSERT_ROWS|номер|количество - вставить несколько строк" & vbCrLf
    prompt = prompt & "INSERT_COLUMN|буква - вставить столбец" & vbCrLf
    prompt = prompt & "INSERT_COLUMNS|буква|количество - вставить несколько столбцов" & vbCrLf
    prompt = prompt & "DELETE_ROW|номер - удалить строку" & vbCrLf
    prompt = prompt & "DELETE_ROWS|начало|конец - удалить строки" & vbCrLf
    prompt = prompt & "DELETE_COLUMN|буква - удалить столбец" & vbCrLf
    prompt = prompt & "DELETE_COLUMNS|начало|конец - удалить столбцы" & vbCrLf
    prompt = prompt & "HIDE_ROW|номер - скрыть строку" & vbCrLf
    prompt = prompt & "HIDE_ROWS|начало|конец - скрыть строки" & vbCrLf
    prompt = prompt & "SHOW_ROW|номер - показать строку" & vbCrLf
    prompt = prompt & "SHOW_ROWS|начало|конец - показать строки" & vbCrLf
    prompt = prompt & "HIDE_COLUMN|буква - скрыть столбец" & vbCrLf
    prompt = prompt & "SHOW_COLUMN|буква - показать столбец" & vbCrLf
    prompt = prompt & "GROUP_ROWS|начало|конец - группировать строки" & vbCrLf
    prompt = prompt & "UNGROUP_ROWS|начало|конец - разгруппировать строки" & vbCrLf
    prompt = prompt & "GROUP_COLUMNS|начало|конец - группировать столбцы" & vbCrLf
    prompt = prompt & "UNGROUP_COLUMNS|начало|конец - разгруппировать столбцы" & vbCrLf & vbCrLf

    ' СОРТИРОВКА И ФИЛЬТРАЦИЯ
    prompt = prompt & "--- СОРТИРОВКА И ФИЛЬТРАЦИЯ ---" & vbCrLf
    prompt = prompt & "SORT|диапазон|колонка|ASC/DESC - сортировка" & vbCrLf
    prompt = prompt & "SORT_MULTI|диапазон|кол1|порядок1|кол2|порядок2 - многоуровневая сортировка" & vbCrLf
    prompt = prompt & "AUTOFILTER|диапазон - включить автофильтр" & vbCrLf
    prompt = prompt & "FILTER|диапазон|колонка|значение - фильтровать" & vbCrLf
    prompt = prompt & "FILTER_TOP|диапазон|колонка|количество - топ N значений" & vbCrLf
    prompt = prompt & "CLEAR_FILTER|диапазон - очистить фильтр" & vbCrLf
    prompt = prompt & "REMOVE_AUTOFILTER - убрать автофильтр" & vbCrLf
    prompt = prompt & "REMOVE_DUPLICATES|диапазон|колонки - удалить дубликаты" & vbCrLf
    prompt = prompt & "FIND_REPLACE|что|на_что - найти и заменить" & vbCrLf
    prompt = prompt & "FIND_REPLACE_RANGE|диапазон|что|на_что - замена в диапазоне" & vbCrLf & vbCrLf

    ' ГРАФИКИ
    prompt = prompt & "--- ГРАФИКИ ---" & vbCrLf
    prompt = prompt & "CREATE_CHART|диапазон|тип|название - создать график" & vbCrLf
    prompt = prompt & "  Типы: LINE, BAR, COLUMN, PIE, AREA, SCATTER, DOUGHNUT" & vbCrLf
    prompt = prompt & "  ВАЖНО: Выбирай ТОЛЬКО нужные столбцы! Используй несмежные диапазоны через запятую." & vbCrLf
    prompt = prompt & "  Примеры:" & vbCrLf
    prompt = prompt & "    CREATE_CHART|A2:A5,B2:B5|LINE|Сумма по датам - ось X=даты, Y=суммы" & vbCrLf
    prompt = prompt & "    CREATE_CHART|A2:B5|COLUMN|Продажи - два столбца (категории + значения)" & vbCrLf
    prompt = prompt & "    CREATE_CHART|B2:B5|PIE|Распределение - один столбец для круговой" & vbCrLf
    prompt = prompt & "CREATE_CHART_AT|диапазон|тип|название|ячейка - график в указанной ячейке" & vbCrLf
    prompt = prompt & "CHART_TITLE|LAST|текст - заголовок" & vbCrLf
    prompt = prompt & "CHART_LEGEND|LAST|позиция - легенда (TOP,BOTTOM,LEFT,RIGHT,NONE)" & vbCrLf
    prompt = prompt & "CHART_AXIS_TITLE|LAST|X|текст - подпись оси X" & vbCrLf
    prompt = prompt & "CHART_AXIS_TITLE|LAST|Y|текст - подпись оси Y" & vbCrLf
    prompt = prompt & "CHART_TYPE|LAST|тип - изменить тип" & vbCrLf
    prompt = prompt & "CHART_MOVE|LAST|ячейка - переместить" & vbCrLf
    prompt = prompt & "CHART_RESIZE|LAST|ширина|высота - размер" & vbCrLf
    prompt = prompt & "CHART_DELETE|LAST - удалить последний" & vbCrLf
    prompt = prompt & "CHART_DELETE_ALL - удалить все" & vbCrLf
    prompt = prompt & "  Индекс: LAST=последний, 1=первый, 2=второй..." & vbCrLf & vbCrLf

    ' СВОДНЫЕ ТАБЛИЦЫ
    prompt = prompt & "--- СВОДНЫЕ ТАБЛИЦЫ ---" & vbCrLf
    prompt = prompt & "CREATE_PIVOT|источник|назначение|имя - создать сводную" & vbCrLf
    prompt = prompt & "PIVOT_ADD_ROW|имя|поле - добавить поле в строки" & vbCrLf
    prompt = prompt & "PIVOT_ADD_COLUMN|имя|поле - добавить поле в столбцы" & vbCrLf
    prompt = prompt & "PIVOT_ADD_VALUE|имя|поле|функция - добавить значение (SUM,COUNT,AVERAGE,MAX,MIN)" & vbCrLf
    prompt = prompt & "PIVOT_ADD_FILTER|имя|поле - добавить фильтр" & vbCrLf
    prompt = prompt & "PIVOT_REFRESH|имя - обновить сводную" & vbCrLf
    prompt = prompt & "PIVOT_REFRESH_ALL - обновить все сводные" & vbCrLf & vbCrLf

    ' ЛИСТЫ
    prompt = prompt & "--- ЛИСТЫ ---" & vbCrLf
    prompt = prompt & "ADD_SHEET|имя - добавить лист" & vbCrLf
    prompt = prompt & "ADD_SHEET_AFTER|имя|после - добавить лист после" & vbCrLf
    prompt = prompt & "DELETE_SHEET|имя - удалить лист" & vbCrLf
    prompt = prompt & "RENAME_SHEET|старое|новое - переименовать" & vbCrLf
    prompt = prompt & "COPY_SHEET|имя|новое_имя - копировать лист" & vbCrLf
    prompt = prompt & "MOVE_SHEET|имя|позиция - переместить лист" & vbCrLf
    prompt = prompt & "HIDE_SHEET|имя - скрыть лист" & vbCrLf
    prompt = prompt & "SHOW_SHEET|имя - показать лист" & vbCrLf
    prompt = prompt & "ACTIVATE_SHEET|имя - активировать лист" & vbCrLf
    prompt = prompt & "TAB_COLOR|имя|цвет - цвет ярлыка" & vbCrLf
    prompt = prompt & "PROTECT_SHEET|имя|пароль - защитить лист" & vbCrLf
    prompt = prompt & "UNPROTECT_SHEET|имя|пароль - снять защиту" & vbCrLf & vbCrLf

    ' ИМЕНОВАННЫЕ ДИАПАЗОНЫ
    prompt = prompt & "--- ИМЕНОВАННЫЕ ДИАПАЗОНЫ ---" & vbCrLf
    prompt = prompt & "CREATE_NAME|имя|диапазон - создать имя" & vbCrLf
    prompt = prompt & "DELETE_NAME|имя - удалить имя" & vbCrLf & vbCrLf

    ' УСЛОВНОЕ ФОРМАТИРОВАНИЕ
    prompt = prompt & "--- УСЛОВНОЕ ФОРМАТИРОВАНИЕ ---" & vbCrLf
    prompt = prompt & "COND_HIGHLIGHT|диапазон|оператор|значение|цвет - подсветка (оператор: >,<,=,>=,<=,<>,BETWEEN)" & vbCrLf
    prompt = prompt & "COND_TOP|диапазон|количество|цвет - топ N" & vbCrLf
    prompt = prompt & "COND_BOTTOM|диапазон|количество|цвет - последние N" & vbCrLf
    prompt = prompt & "COND_DUPLICATE|диапазон|цвет - дубликаты" & vbCrLf
    prompt = prompt & "COND_UNIQUE|диапазон|цвет - уникальные" & vbCrLf
    prompt = prompt & "COND_TEXT|диапазон|текст|цвет - содержит текст" & vbCrLf
    prompt = prompt & "COND_BLANK|диапазон|цвет - пустые ячейки" & vbCrLf
    prompt = prompt & "COND_NOT_BLANK|диапазон|цвет - непустые ячейки" & vbCrLf
    prompt = prompt & "DATA_BARS|диапазон|цвет - гистограммы" & vbCrLf
    prompt = prompt & "COLOR_SCALE|диапазон|цвет1|цвет2 - цветовая шкала" & vbCrLf
    prompt = prompt & "COLOR_SCALE3|диапазон|цвет1|цвет2|цвет3 - 3-цветная шкала" & vbCrLf
    prompt = prompt & "ICON_SET|диапазон|набор - значки (ARROWS,FLAGS,STARS,BARS)" & vbCrLf
    prompt = prompt & "CLEAR_COND_FORMAT|диапазон - очистить условное форматирование" & vbCrLf & vbCrLf

    ' ПРОВЕРКА ДАННЫХ
    prompt = prompt & "--- ПРОВЕРКА ДАННЫХ ---" & vbCrLf
    prompt = prompt & "VALIDATION_LIST|диапазон|значения - выпадающий список (значения через ;)" & vbCrLf
    prompt = prompt & "VALIDATION_NUMBER|диапазон|мин|макс - числа в диапазоне" & vbCrLf
    prompt = prompt & "VALIDATION_DATE|диапазон|начало|конец - даты в диапазоне" & vbCrLf
    prompt = prompt & "VALIDATION_TEXT_LENGTH|диапазон|мин|макс - длина текста" & vbCrLf
    prompt = prompt & "VALIDATION_CUSTOM|диапазон|формула - произвольная формула" & vbCrLf
    prompt = prompt & "CLEAR_VALIDATION|диапазон - очистить проверку" & vbCrLf & vbCrLf

    ' КОММЕНТАРИИ И ПРИМЕЧАНИЯ
    prompt = prompt & "--- КОММЕНТАРИИ ---" & vbCrLf
    prompt = prompt & "ADD_COMMENT|адрес|текст - добавить комментарий" & vbCrLf
    prompt = prompt & "EDIT_COMMENT|адрес|текст - изменить комментарий" & vbCrLf
    prompt = prompt & "DELETE_COMMENT|адрес - удалить комментарий" & vbCrLf
    prompt = prompt & "SHOW_COMMENT|адрес - показать комментарий" & vbCrLf
    prompt = prompt & "HIDE_COMMENT|адрес - скрыть комментарий" & vbCrLf
    prompt = prompt & "SHOW_ALL_COMMENTS - показать все" & vbCrLf
    prompt = prompt & "HIDE_ALL_COMMENTS - скрыть все" & vbCrLf & vbCrLf

    ' ГИПЕРССЫЛКИ
    prompt = prompt & "--- ГИПЕРССЫЛКИ ---" & vbCrLf
    prompt = prompt & "ADD_HYPERLINK|адрес|url|текст - добавить ссылку" & vbCrLf
    prompt = prompt & "ADD_HYPERLINK_CELL|адрес|ссылка_на_ячейку|текст - ссылка на ячейку" & vbCrLf
    prompt = prompt & "REMOVE_HYPERLINK|адрес - удалить ссылку" & vbCrLf & vbCrLf

    ' ЗАЩИТА
    prompt = prompt & "--- ЗАЩИТА ---" & vbCrLf
    prompt = prompt & "LOCK_CELLS|диапазон - заблокировать ячейки" & vbCrLf
    prompt = prompt & "UNLOCK_CELLS|диапазон - разблокировать ячейки" & vbCrLf & vbCrLf

    ' ОБЛАСТЬ ПРОСМОТРА
    prompt = prompt & "--- ОБЛАСТЬ ПРОСМОТРА ---" & vbCrLf
    prompt = prompt & "FREEZE_PANES|адрес - закрепить области" & vbCrLf
    prompt = prompt & "FREEZE_TOP_ROW - закрепить верхнюю строку" & vbCrLf
    prompt = prompt & "FREEZE_FIRST_COLUMN - закрепить первый столбец" & vbCrLf
    prompt = prompt & "UNFREEZE_PANES - снять закрепление" & vbCrLf
    prompt = prompt & "ZOOM|процент - масштаб" & vbCrLf
    prompt = prompt & "GOTO|адрес - перейти к ячейке" & vbCrLf
    prompt = prompt & "SELECT|диапазон - выделить диапазон" & vbCrLf & vbCrLf

    ' ПЕЧАТЬ
    prompt = prompt & "--- ПЕЧАТЬ ---" & vbCrLf
    prompt = prompt & "SET_PRINT_AREA|диапазон - область печати" & vbCrLf
    prompt = prompt & "CLEAR_PRINT_AREA - очистить область печати" & vbCrLf
    prompt = prompt & "PAGE_ORIENTATION|PORTRAIT/LANDSCAPE - ориентация" & vbCrLf
    prompt = prompt & "PAGE_MARGINS|лево|право|верх|низ - поля (в см)" & vbCrLf
    prompt = prompt & "PRINT_TITLES_ROWS|начало|конец - печатать строки заголовков" & vbCrLf
    prompt = prompt & "PRINT_TITLES_COLS|начало|конец - печатать столбцы заголовков" & vbCrLf
    prompt = prompt & "PRINT_GRIDLINES|TRUE/FALSE - печатать сетку" & vbCrLf
    prompt = prompt & "FIT_TO_PAGE|ширина|высота - вписать на страницы" & vbCrLf & vbCrLf

    ' ИЗОБРАЖЕНИЯ
    prompt = prompt & "--- ИЗОБРАЖЕНИЯ ---" & vbCrLf
    prompt = prompt & "INSERT_PICTURE|путь|лево|верх|ширина|высота - вставить изображение" & vbCrLf
    prompt = prompt & "DELETE_PICTURES - удалить все изображения" & vbCrLf & vbCrLf

    ' ФОРМЫ
    prompt = prompt & "--- ФОРМЫ ---" & vbCrLf
    prompt = prompt & "ADD_BUTTON|лево|верх|ширина|высота|текст - добавить кнопку" & vbCrLf
    prompt = prompt & "ADD_CHECKBOX|адрес|текст - добавить флажок" & vbCrLf
    prompt = prompt & "ADD_DROPDOWN|адрес|значения - добавить выпадающий список" & vbCrLf
    prompt = prompt & "DELETE_SHAPES - удалить все фигуры" & vbCrLf & vbCrLf

    ' СПЕЦИАЛЬНЫЕ
    prompt = prompt & "--- СПЕЦИАЛЬНЫЕ ---" & vbCrLf
    prompt = prompt & "CALCULATE - пересчитать" & vbCrLf
    prompt = prompt & "CALCULATE_SHEET - пересчитать лист" & vbCrLf
    prompt = prompt & "TEXT_TO_COLUMNS|диапазон|разделитель - текст по столбцам" & vbCrLf
    prompt = prompt & "REMOVE_SPACES|диапазон - удалить лишние пробелы" & vbCrLf
    prompt = prompt & "UPPER_CASE|диапазон - в ВЕРХНИЙ РЕГИСТР" & vbCrLf
    prompt = prompt & "LOWER_CASE|диапазон - в нижний регистр" & vbCrLf
    prompt = prompt & "PROPER_CASE|диапазон - Каждое Слово С Заглавной" & vbCrLf
    prompt = prompt & "FLASH_FILL|диапазон - мгновенное заполнение" & vbCrLf
    prompt = prompt & "SUBTOTAL|диапазон|функция|колонка - промежуточные итоги (SUM,COUNT,AVERAGE)" & vbCrLf
    prompt = prompt & "REMOVE_SUBTOTALS - удалить промежуточные итоги" & vbCrLf & vbCrLf

    prompt = prompt & "=== КОНЕЦ СПИСКА КОМАНД ===" & vbCrLf & vbCrLf
    prompt = prompt & "Отвечай на русском. ВСЕГДА включай блок ```commands``` с командами!" & vbCrLf
    prompt = prompt & "Если задача сложная — используй несколько команд последовательно." & vbCrLf & vbCrLf

    If Len(excelContext) > 0 Then
        prompt = prompt & "Контекст Excel:" & vbCrLf & excelContext
    End If

    BuildSystemPrompt = prompt
End Function


'----------------------------------------
' Построение JSON запроса
'----------------------------------------
' ---
' Что делает: Собирает служебную строку или JSON для следующего шага.
' Вход: systemPrompt, userMessage, modelName
' Выход: значение функции
' ---
Private Function BuildRequestJSON(systemPrompt As String, userMessage As String, modelName As String) As String
    Dim json As String

    ' Экранируем специальные символы
    systemPrompt = EscapeJSON(systemPrompt)
    userMessage = EscapeJSON(userMessage)

    json = "{"
    json = json & """model"": """ & modelName & ""","
    json = json & """messages"": ["
    json = json & "{""role"": ""system"", ""content"": """ & systemPrompt & """},"
    json = json & "{""role"": ""user"", ""content"": """ & userMessage & """}"
    json = json & "],"
    json = json & """temperature"": 0.1,"
    json = json & """max_tokens"": 4096"
    json = json & "}"

    BuildRequestJSON = json
End Function


'----------------------------------------
' Экранирование JSON
'----------------------------------------
' ---
' Что делает: Вспомогательная функция модуля.
' Вход: text
' Выход: значение функции
' ---
Private Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSON = result
End Function


'----------------------------------------
' HTTP-запрос
'----------------------------------------
' ---
' Что делает: Отправляет запрос во внешний сервис и получает ответ.
' Вход: url, apiKey, body, model
' Выход: значение функции
' ---
Private Function SendHTTPRequest(url As String, apiKey As String, body As String, model As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim httpCreated As Boolean

    httpCreated = False

    ' Проверяем наличие ключа
    If Len(Trim(apiKey)) = 0 Then
        SendHTTPRequest = "{""error"": ""API-ключ не настроен. Откройте Настройки и введите ключ.""}"
        Exit Function
    End If

    ' Создаём HTTP-объект (предпочитаем ServerXMLHTTP для таймаутов)
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.ServerXMLHTTP")
    End If
    If http Is Nothing Then
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    End If
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.XMLHTTP")
    End If
    On Error GoTo ErrorHandler

    If http Is Nothing Then
        SendHTTPRequest = "{""error"": ""Не удалось создать HTTP-объект.""}"
        Exit Function
    End If

    httpCreated = True

    ' Устанавливаем таймауты (если поддерживается)
    On Error Resume Next
    ' setTimeouts: resolve, connect, send, receive (в миллисекундах)
    http.setTimeouts 5000, 10000, 60000, 120000 ' 5с, 10с, 60с, 120с
    On Error GoTo ErrorHandler

    ' Открываем соединение
    http.Open "POST", url, False

    ' Устанавливаем заголовки
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey

    ' Дополнительные заголовки для OpenRouter
    If model = "claude" Then
        http.setRequestHeader "HTTP-Referer", "https://excel-ai-assistant.local"
        http.setRequestHeader "X-Title", "Excel AI Assistant VBA"
    End If

    ' Отправляем запрос
    http.send body

    ' Проверяем результат
    If http.Status = 200 Then
        SendHTTPRequest = http.responseText
    ElseIf http.Status = 401 Then
        SendHTTPRequest = "{""error"": ""Ошибка авторизации (401). Проверьте API-ключ.""}"
    ElseIf http.Status = 403 Then
        SendHTTPRequest = "{""error"": ""Доступ запрещён (403). Проверьте API-ключ и лимиты.""}"
    ElseIf http.Status = 429 Then
        SendHTTPRequest = "{""error"": ""Превышен лимит запросов (429). Подождите.""}"
    ElseIf http.Status >= 500 Then
        SendHTTPRequest = "{""error"": ""Ошибка сервера (" & http.Status & ").""}"
    Else
        SendHTTPRequest = "{""error"": ""HTTP " & http.Status & ": " & http.statusText & """}"
    End If

    Set http = Nothing
    Exit Function

ErrorHandler:
    Dim errMsg As String
    errMsg = "Ошибка " & Err.Number & ": " & Err.Description
    SendHTTPRequest = "{""error"": """ & EscapeJSON(errMsg) & """}"
    If httpCreated Then Set http = Nothing
End Function


'----------------------------------------
' Парсинг ответа JSON
'----------------------------------------
' ---
' Что делает: Разбирает текст и извлекает структурированные данные.
' Вход: jsonResponse
' Выход: значение функции
' ---
Private Function ParseResponse(jsonResponse As String) As String
    On Error GoTo ErrorHandler

    Dim content As String
    Dim startPos As Long
    Dim endPos As Long

    ' Проверяем на ошибку
    If InStr(jsonResponse, """error""") > 0 Then
        startPos = InStr(jsonResponse, """error""") + 10
        endPos = InStr(startPos, jsonResponse, """")
        If endPos > startPos Then
            ParseResponse = "ОШИБКА API: " & Mid(jsonResponse, startPos, endPos - startPos)
        Else
            ParseResponse = "ОШИБКА API: Неизвестная ошибка"
        End If
        Exit Function
    End If

    ' Ищем content в ответе
    startPos = InStr(jsonResponse, """content"":")
    If startPos = 0 Then
        ParseResponse = "ОШИБКА: Не удалось разобрать ответ"
        Exit Function
    End If

    ' Находим начало значения
    startPos = InStr(startPos, jsonResponse, ":") + 1

    ' Пропускаем пробелы
    Do While Mid(jsonResponse, startPos, 1) = " "
        startPos = startPos + 1
    Loop

    ' Проверяем, начинается ли с кавычки
    If Mid(jsonResponse, startPos, 1) = """" Then
        startPos = startPos + 1
        ' Ищем закрывающую кавычку (не экранированную)
        endPos = startPos
        Do
            endPos = InStr(endPos, jsonResponse, """")
            If endPos = 0 Then Exit Do
            ' Проверяем, не экранирована ли
            If Mid(jsonResponse, endPos - 1, 1) <> "\" Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop

        If endPos > startPos Then
            content = Mid(jsonResponse, startPos, endPos - startPos)
        End If
    End If

    ' Убираем экранирование
    content = Replace(content, "\n", vbCrLf)
    content = Replace(content, "\t", vbTab)
    content = Replace(content, "\""", """")
    content = Replace(content, "\\", "\")

    ParseResponse = content
    Exit Function

ErrorHandler:
    ParseResponse = "ОШИБКА парсинга: " & Err.Description
End Function


'----------------------------------------
' Проверка доступности LM Studio
'----------------------------------------
' ---
' Что делает: Проверяет условие и возвращает True/False.
' Вход: нет
' Выход: значение функции
' ---
Public Function IsLMStudioAvailable() As Boolean
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim ip As String
    Dim port As String

    ip = GetLMStudioSetting("IP")
    port = GetLMStudioSetting("Port")
    url = "http://" & ip & ":" & port & "/v1/models"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    End If

    http.setTimeouts 2000, 2000, 2000, 2000
    http.Open "GET", url, False
    http.send

    IsLMStudioAvailable = (http.Status = 200)
    Set http = Nothing
    Exit Function

ErrorHandler:
    IsLMStudioAvailable = False
End Function


'----------------------------------------
' Получение списка моделей из LM Studio
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: нет
' Выход: значение функции
' ---
Public Function GetLMStudioModels() As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim ip As String
    Dim port As String
    Dim response As String
    Dim models As String
    Dim pos As Long
    Dim endPos As Long
    Dim modelId As String

    ip = GetLMStudioSetting("IP")
    port = GetLMStudioSetting("Port")
    url = "http://" & ip & ":" & port & "/v1/models"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    End If

    http.setTimeouts 5000, 5000, 5000, 5000
    http.Open "GET", url, False
    http.send

    If http.Status <> 200 Then
        GetLMStudioModels = "ERROR:HTTP " & http.Status
        Set http = Nothing
        Exit Function
    End If

    response = http.responseText
    Set http = Nothing

    ' Парсим JSON для извлечения id моделей
    models = ""
    pos = 1
    Do
        pos = InStr(pos, response, """id""")
        If pos = 0 Then Exit Do

        pos = InStr(pos, response, ":")
        If pos = 0 Then Exit Do

        pos = InStr(pos, response, """")
        If pos = 0 Then Exit Do
        pos = pos + 1

        endPos = InStr(pos, response, """")
        If endPos = 0 Then Exit Do

        modelId = Mid(response, pos, endPos - pos)

        If Len(models) > 0 Then models = models & "|"
        models = models & modelId

        pos = endPos + 1
    Loop

    GetLMStudioModels = models
    Exit Function

ErrorHandler:
    GetLMStudioModels = "ERROR:" & Err.Description
End Function


'----------------------------------------
' Отправка запроса к локальной модели LM Studio
'----------------------------------------
' ---
' Что делает: Отправляет запрос во внешний сервис и получает ответ.
' Вход: userMessage, excelContext
' Выход: значение функции
' ---
Public Function SendToLocalAI(userMessage As String, Optional excelContext As String = "") As String
    On Error GoTo ErrorHandler

    Dim ip As String
    Dim port As String
    Dim modelName As String
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim systemPrompt As String

    ip = GetLMStudioSetting("IP")
    port = GetLMStudioSetting("Port")
    modelName = GetLMStudioSetting("Model")

    If Len(ip) = 0 Or Len(port) = 0 Then
        SendToLocalAI = "ОШИБКА: Настройки LM Studio не заданы. Откройте настройки."
        Exit Function
    End If

    url = "http://" & ip & ":" & port & "/v1/chat/completions"

    ' Если модель не указана - пробуем получить первую доступную
    If Len(modelName) = 0 Then
        Dim models As String
        models = GetLMStudioModels()
        If Left(models, 6) = "ERROR:" Then
            SendToLocalAI = "ОШИБКА: Не удалось получить список моделей: " & Mid(models, 7)
            Exit Function
        End If
        If Len(models) = 0 Then
            SendToLocalAI = "ОШИБКА: В LM Studio нет загруженных моделей."
            Exit Function
        End If
        ' Берём первую модель
        If InStr(models, "|") > 0 Then
            modelName = Left(models, InStr(models, "|") - 1)
        Else
            modelName = models
        End If
    End If

    ' Формируем системный промпт
    systemPrompt = BuildSystemPrompt(excelContext)

    ' Формируем JSON запрос
    requestBody = BuildRequestJSON(systemPrompt, userMessage, modelName)

    ' Отправляем запрос
    response = SendLocalHTTPRequest(url, requestBody)

    ' Парсим ответ
    SendToLocalAI = ParseResponse(response)
    Exit Function

ErrorHandler:
    SendToLocalAI = "ОШИБКА: " & Err.Description
End Function


'----------------------------------------
' HTTP-запрос для локальной модели (без API-ключа)
'----------------------------------------
' ---
' Что делает: Отправляет запрос во внешний сервис и получает ответ.
' Вход: url, body
' Выход: значение функции
' ---
Private Function SendLocalHTTPRequest(url As String, body As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object

    ' Создаём HTTP-объект
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.ServerXMLHTTP")
    End If
    If http Is Nothing Then
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    End If
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.XMLHTTP")
    End If
    On Error GoTo ErrorHandler

    If http Is Nothing Then
        SendLocalHTTPRequest = "{""error"": ""Не удалось создать HTTP-объект.""}"
        Exit Function
    End If

    ' Устанавливаем таймауты
    On Error Resume Next
    http.setTimeouts 5000, 10000, 120000, 300000 ' 5с, 10с, 120с, 300с
    On Error GoTo ErrorHandler

    ' Открываем соединение
    http.Open "POST", url, False

    ' Устанавливаем заголовки
    http.setRequestHeader "Content-Type", "application/json"
    ' LM Studio не требует API-ключ, но можно указать любой
    http.setRequestHeader "Authorization", "Bearer lm-studio"

    ' Отправляем запрос
    http.send body

    ' Получаем ответ
    If http.Status = 200 Then
        SendLocalHTTPRequest = http.responseText
    Else
        SendLocalHTTPRequest = "{""error"": ""HTTP " & http.Status & ": " & http.statusText & """}"
    End If

    Set http = Nothing
    Exit Function

ErrorHandler:
    If Not http Is Nothing Then Set http = Nothing
    SendLocalHTTPRequest = "{""error"": """ & EscapeJSON(Err.Description) & """}"
End Function
