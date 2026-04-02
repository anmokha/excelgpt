' =============================================================================
' modAICommands.bas
' Что внутри: Парсинг и выполнение AI-команд в Excel (формат commands).
' Автоматически собран из модульной структуры проекта.
' =============================================================================

Option Explicit


'----------------------------------------
' Извлечение команд из ответа
'----------------------------------------
' ---
' Что делает: Разбирает текст и извлекает структурированные данные.
' Вход: response
' Выход: значение функции
' ---
Public Function ExtractCommands(response As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim commands As String

    startPos = InStr(response, "```commands")
    If startPos = 0 Then
        ExtractCommands = ""
        Exit Function
    End If

    startPos = startPos + Len("```commands") + 1
    endPos = InStr(startPos, response, "```")

    If endPos = 0 Then
        ExtractCommands = ""
        Exit Function
    End If

    commands = Trim(Mid(response, startPos, endPos - startPos))
    ExtractCommands = commands
End Function


'----------------------------------------
' Выполнение команд
'----------------------------------------
' ---
' Что делает: Выполняет подготовленные команды в Excel.
' Вход: commands
' Выход: значение функции
' ---
Public Function ExecuteCommands(commands As String) As String
    On Error GoTo ErrorHandler

    Dim lines() As String
    Dim i As Long
    Dim cmd As String
    Dim executedCount As Long
    Dim result As String

    If Len(commands) = 0 Then
        ExecuteCommands = ""
        Exit Function
    End If

    lines = Split(commands, vbLf)
    executedCount = 0

    For i = 0 To UBound(lines)
        cmd = Trim(Replace(lines(i), vbCr, ""))
        If Len(cmd) > 0 Then
            If ExecuteSingleCommand(cmd) Then
                executedCount = executedCount + 1
            End If
        End If
    Next i

    result = "[Выполнено команд: " & executedCount & "]"
    ExecuteCommands = result
    Exit Function

ErrorHandler:
    ExecuteCommands = "Ошибка выполнения: " & Err.Description
End Function



'----------------------------------------
' Парсинг цвета
'----------------------------------------
' ---
' Что делает: Разбирает текст и извлекает структурированные данные.
' Вход: colorStr
' Выход: значение функции
' ---
Private Function ParseColor(colorStr As String) As Long
    On Error Resume Next

    Dim c As String
    Dim rgbParts() As String

    c = UCase(Trim(colorStr))
    Debug.Print "ParseColor: input=[" & colorStr & "] upper=[" & c & "]"

    ' Проверяем RGB формат
    If Len(c) >= 4 And Left(c, 4) = "RGB:" Then
        rgbParts = Split(Mid(c, 5), ",")
        If UBound(rgbParts) >= 2 Then
            ParseColor = RGB(CLng(Trim(rgbParts(0))), CLng(Trim(rgbParts(1))), CLng(Trim(rgbParts(2))))
            Debug.Print "ParseColor: RGB result=" & ParseColor
            Exit Function
        End If
    End If

    ' Предопределённые цвета
    Select Case c
        Case "RED": ParseColor = RGB(255, 0, 0)
        Case "GREEN": ParseColor = RGB(0, 128, 0)
        Case "BLUE": ParseColor = RGB(0, 0, 255)
        Case "YELLOW": ParseColor = RGB(255, 255, 0)
        Case "ORANGE": ParseColor = RGB(255, 165, 0)
        Case "PURPLE": ParseColor = RGB(128, 0, 128)
        Case "PINK": ParseColor = RGB(255, 192, 203)
        Case "CYAN": ParseColor = RGB(0, 255, 255)
        Case "WHITE": ParseColor = RGB(255, 255, 255)
        Case "BLACK": ParseColor = RGB(0, 0, 0)
        Case "GRAY", "GREY": ParseColor = RGB(128, 128, 128)
        Case "LIGHTGRAY", "LIGHTGREY": ParseColor = RGB(192, 192, 192)
        Case "DARKGRAY", "DARKGREY": ParseColor = RGB(64, 64, 64)
        Case "BROWN": ParseColor = RGB(139, 69, 19)
        Case "LIME": ParseColor = RGB(0, 255, 0)
        Case "NAVY": ParseColor = RGB(0, 0, 128)
        Case "TEAL": ParseColor = RGB(0, 128, 128)
        Case "MAROON": ParseColor = RGB(128, 0, 0)
        Case "OLIVE": ParseColor = RGB(128, 128, 0)
        Case "GOLD": ParseColor = RGB(255, 215, 0)
        Case "SILVER": ParseColor = RGB(192, 192, 192)
        Case Else: ParseColor = RGB(0, 0, 0) ' По умолчанию чёрный
    End Select

    Debug.Print "ParseColor: result=" & ParseColor
    If Err.Number <> 0 Then
        Debug.Print "ParseColor: ERROR " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
End Function


'----------------------------------------
' Локализация формулы (английские функции -> русские)
' ПОЛНЫЙ СПИСОК ВСЕХ ФУНКЦИЙ EXCEL
'----------------------------------------
' ---
' Что делает: Вспомогательная функция модуля.
' Вход: formula
' Выход: значение функции
' ---
Private Function LocalizeFormula(formula As String) As String
    Dim result As String
    result = formula

    ' === МАТЕМАТИЧЕСКИЕ ===
    result = ReplaceFunc(result, "ABS", "ABS")
    result = ReplaceFunc(result, "ACOS", "ACOS")
    result = ReplaceFunc(result, "ACOSH", "ACOSH")
    result = ReplaceFunc(result, "ACOT", "ACOT")
    result = ReplaceFunc(result, "ACOTH", "ACOTH")
    result = ReplaceFunc(result, "AGGREGATE", "АГРЕГАТ")
    result = ReplaceFunc(result, "ARABIC", "АРАБСКОЕ")
    result = ReplaceFunc(result, "ASIN", "ASIN")
    result = ReplaceFunc(result, "ASINH", "ASINH")
    result = ReplaceFunc(result, "ATAN", "ATAN")
    result = ReplaceFunc(result, "ATAN2", "ATAN2")
    result = ReplaceFunc(result, "ATANH", "ATANH")
    result = ReplaceFunc(result, "BASE", "ОСНОВАНИЕ")
    result = ReplaceFunc(result, "CEILING.MATH", "ОКРВВЕРХ.МАТ")
    result = ReplaceFunc(result, "CEILING.PRECISE", "ОКРВВЕРХ.ТОЧН")
    result = ReplaceFunc(result, "CEILING", "ОКРВВЕРХ")
    result = ReplaceFunc(result, "COMBIN", "ЧИСЛКОМБ")
    result = ReplaceFunc(result, "COMBINA", "ЧИСЛКОМБА")
    result = ReplaceFunc(result, "COS", "COS")
    result = ReplaceFunc(result, "COSH", "COSH")
    result = ReplaceFunc(result, "COT", "COT")
    result = ReplaceFunc(result, "COTH", "COTH")
    result = ReplaceFunc(result, "CSC", "CSC")
    result = ReplaceFunc(result, "CSCH", "CSCH")
    result = ReplaceFunc(result, "DECIMAL", "ДЕС")
    result = ReplaceFunc(result, "DEGREES", "ГРАДУСЫ")
    result = ReplaceFunc(result, "EVEN", "ЧЁТН")
    result = ReplaceFunc(result, "EXP", "EXP")
    result = ReplaceFunc(result, "FACT", "ФАКТР")
    result = ReplaceFunc(result, "FACTDOUBLE", "ДВФАКТР")
    result = ReplaceFunc(result, "FLOOR.MATH", "ОКРВНИЗ.МАТ")
    result = ReplaceFunc(result, "FLOOR.PRECISE", "ОКРВНИЗ.ТОЧН")
    result = ReplaceFunc(result, "FLOOR", "ОКРВНИЗ")
    result = ReplaceFunc(result, "GCD", "НОД")
    result = ReplaceFunc(result, "INT", "ЦЕЛОЕ")
    result = ReplaceFunc(result, "ISO.CEILING", "ISO.ОКРВВЕРХ")
    result = ReplaceFunc(result, "LCM", "НОК")
    result = ReplaceFunc(result, "LN", "LN")
    result = ReplaceFunc(result, "LOG10", "LOG10")
    result = ReplaceFunc(result, "LOG", "LOG")
    result = ReplaceFunc(result, "MDETERM", "МОПРЕД")
    result = ReplaceFunc(result, "MINVERSE", "МОБР")
    result = ReplaceFunc(result, "MMULT", "МУМНОЖ")
    result = ReplaceFunc(result, "MOD", "ОСТАТ")
    result = ReplaceFunc(result, "MROUND", "ОКРУГЛТ")
    result = ReplaceFunc(result, "MULTINOMIAL", "МУЛЬТИНОМ")
    result = ReplaceFunc(result, "MUNIT", "МЕДИН")
    result = ReplaceFunc(result, "ODD", "НЕЧЁТ")
    result = ReplaceFunc(result, "PI", "ПИ")
    result = ReplaceFunc(result, "POWER", "СТЕПЕНЬ")
    result = ReplaceFunc(result, "PRODUCT", "ПРОИЗВЕД")
    result = ReplaceFunc(result, "QUOTIENT", "ЧАСТНОЕ")
    result = ReplaceFunc(result, "RADIANS", "РАДИАНЫ")
    result = ReplaceFunc(result, "RANDBETWEEN", "СЛУЧМЕЖДУ")
    result = ReplaceFunc(result, "RAND", "СЛЧИС")
    result = ReplaceFunc(result, "ROMAN", "РИМСКОЕ")
    result = ReplaceFunc(result, "ROUNDDOWN", "ОКРУГЛВНИЗ")
    result = ReplaceFunc(result, "ROUNDUP", "ОКРУГЛВВЕРХ")
    result = ReplaceFunc(result, "ROUND", "ОКРУГЛ")
    result = ReplaceFunc(result, "SEC", "SEC")
    result = ReplaceFunc(result, "SECH", "SECH")
    result = ReplaceFunc(result, "SERIESSUM", "РЯД.СУММ")
    result = ReplaceFunc(result, "SIGN", "ЗНАК")
    result = ReplaceFunc(result, "SIN", "SIN")
    result = ReplaceFunc(result, "SINH", "SINH")
    result = ReplaceFunc(result, "SQRT", "КОРЕНЬ")
    result = ReplaceFunc(result, "SQRTPI", "КОРЕ|НЬ.ПИ")
    result = ReplaceFunc(result, "SUBTOTAL", "ПРОМЕЖУТОЧНЫЕ.ИТОГИ")
    result = ReplaceFunc(result, "SUMIFS", "СУММЕСЛИМН")
    result = ReplaceFunc(result, "SUMIF", "СУММЕСЛИ")
    result = ReplaceFunc(result, "SUMPRODUCT", "СУММПРОИЗВ")
    result = ReplaceFunc(result, "SUMSQ", "СУММКВ")
    result = ReplaceFunc(result, "SUMX2MY2", "СУММРАЗНКВ")
    result = ReplaceFunc(result, "SUMX2PY2", "СУММСУММКВ")
    result = ReplaceFunc(result, "SUMXMY2", "СУММКВРАЗН")
    result = ReplaceFunc(result, "SUM", "СУММ")
    result = ReplaceFunc(result, "TAN", "TAN")
    result = ReplaceFunc(result, "TANH", "TANH")
    result = ReplaceFunc(result, "TRUNC", "ОТБР")

    ' === ЛОГИЧЕСКИЕ ===
    result = ReplaceFunc(result, "AND", "И")
    result = ReplaceFunc(result, "FALSE", "ЛОЖЬ")
    result = ReplaceFunc(result, "IFERROR", "ЕСЛИОШИБКА")
    result = ReplaceFunc(result, "IFNA", "ЕСНД")
    result = ReplaceFunc(result, "IFS", "УСЛОВИЯ")
    result = ReplaceFunc(result, "IF", "ЕСЛИ")
    result = ReplaceFunc(result, "NOT", "НЕ")
    result = ReplaceFunc(result, "OR", "ИЛИ")
    result = ReplaceFunc(result, "SWITCH", "ПЕРЕКЛЮЧ")
    result = ReplaceFunc(result, "TRUE", "ИСТИНА")
    result = ReplaceFunc(result, "XOR", "ИСКЛИЛИ")

    ' === ТЕКСТОВЫЕ ===
    result = ReplaceFunc(result, "ASC", "ASC")
    result = ReplaceFunc(result, "BAHTTEXT", "БАТТ.ТЕКСТ")
    result = ReplaceFunc(result, "CHAR", "СИМВОЛ")
    result = ReplaceFunc(result, "CLEAN", "ПЕЧСИМВ")
    result = ReplaceFunc(result, "CODE", "КОДСИМВ")
    result = ReplaceFunc(result, "CONCATENATE", "СЦЕПИТЬ")
    result = ReplaceFunc(result, "CONCAT", "СЦЕП")
    result = ReplaceFunc(result, "DOLLAR", "РУБЛЬ")
    result = ReplaceFunc(result, "EXACT", "СОВПАД")
    result = ReplaceFunc(result, "FIND", "НАЙТИ")
    result = ReplaceFunc(result, "FIXED", "ФИКСИРОВАННЫЙ")
    result = ReplaceFunc(result, "LEFT", "ЛЕВСИМВ")
    result = ReplaceFunc(result, "LEN", "ДЛСТР")
    result = ReplaceFunc(result, "LOWER", "СТРОЧН")
    result = ReplaceFunc(result, "MID", "ПСТР")
    result = ReplaceFunc(result, "NUMBERVALUE", "ЧЗНАЧ")
    result = ReplaceFunc(result, "PHONETIC", "ФОНЕТИЧЕСКАЯ")
    result = ReplaceFunc(result, "PROPER", "ПРОПНАЧ")
    result = ReplaceFunc(result, "REPLACE", "ЗАМЕНИТЬ")
    result = ReplaceFunc(result, "REPT", "ПОВТОР")
    result = ReplaceFunc(result, "RIGHT", "ПРАВСИМВ")
    result = ReplaceFunc(result, "SEARCH", "ПОИСК")
    result = ReplaceFunc(result, "SUBSTITUTE", "ПОДСТАВИТЬ")
    result = ReplaceFunc(result, "TEXTJOIN", "ОБЪЕДИНИТЬ")
    result = ReplaceFunc(result, "TEXT", "ТЕКСТ")
    result = ReplaceFunc(result, "TRIM", "СЖПРОБЕЛЫ")
    result = ReplaceFunc(result, "UNICHAR", "ЮНИСИМВ")
    result = ReplaceFunc(result, "UNICODE", "ЮНИКОД")
    result = ReplaceFunc(result, "UPPER", "ПРОПИСН")
    result = ReplaceFunc(result, "VALUE", "ЗНАЧЕН")

    ' === ДАТА И ВРЕМЯ ===
    result = ReplaceFunc(result, "DATE", "ДАТА")
    result = ReplaceFunc(result, "DATEDIF", "РАЗНДАТ")
    result = ReplaceFunc(result, "DATEVALUE", "ДАТАЗНАЧ")
    result = ReplaceFunc(result, "DAY", "ДЕНЬ")
    result = ReplaceFunc(result, "DAYS360", "ДНЕЙ360")
    result = ReplaceFunc(result, "DAYS", "ДНИ")
    result = ReplaceFunc(result, "EDATE", "ДАТАМЕС")
    result = ReplaceFunc(result, "EOMONTH", "КОНМЕСЯЦА")
    result = ReplaceFunc(result, "HOUR", "ЧАС")
    result = ReplaceFunc(result, "ISOWEEKNUM", "НОМНЕДЕЛИ.ISO")
    result = ReplaceFunc(result, "MINUTE", "МИНУТЫ")
    result = ReplaceFunc(result, "MONTH", "МЕСЯЦ")
    result = ReplaceFunc(result, "NETWORKDAYS.INTL", "ЧИСТРАБДНИ.МЕЖД")
    result = ReplaceFunc(result, "NETWORKDAYS", "ЧИСТРАБДНИ")
    result = ReplaceFunc(result, "NOW", "ТДАТА")
    result = ReplaceFunc(result, "SECOND", "СЕКУНДЫ")
    result = ReplaceFunc(result, "TIMEVALUE", "ВРЕМЗНАЧ")
    result = ReplaceFunc(result, "TIME", "ВРЕМЯ")
    result = ReplaceFunc(result, "TODAY", "СЕГОДНЯ")
    result = ReplaceFunc(result, "WEEKDAY", "ДЕНЬНЕД")
    result = ReplaceFunc(result, "WEEKNUM", "НОМНЕДЕЛИ")
    result = ReplaceFunc(result, "WORKDAY.INTL", "РАБДЕНЬ.МЕЖД")
    result = ReplaceFunc(result, "WORKDAY", "РАБДЕНЬ")
    result = ReplaceFunc(result, "YEARFRAC", "ДОЛЯГОДА")
    result = ReplaceFunc(result, "YEAR", "ГОД")

    ' === ССЫЛКИ И ПОИСК ===
    result = ReplaceFunc(result, "ADDRESS", "АДРЕС")
    result = ReplaceFunc(result, "AREAS", "ОБЛАСТИ")
    result = ReplaceFunc(result, "CHOOSE", "ВЫБОР")
    result = ReplaceFunc(result, "COLUMNS", "ЧИСЛСТОЛБ")
    result = ReplaceFunc(result, "COLUMN", "СТОЛБЕЦ")
    result = ReplaceFunc(result, "FORMULATEXT", "Ф.ТЕКСТ")
    result = ReplaceFunc(result, "GETPIVOTDATA", "ПОЛУЧИТЬ.ДАННЫЕ.СВОДНОЙ.ТАБЛИЦЫ")
    result = ReplaceFunc(result, "HLOOKUP", "ГПР")
    result = ReplaceFunc(result, "HYPERLINK", "ГИПЕРССЫЛКА")
    result = ReplaceFunc(result, "INDEX", "ИНДЕКС")
    result = ReplaceFunc(result, "INDIRECT", "ДВССЫЛ")
    result = ReplaceFunc(result, "LOOKUP", "ПРОСМОТР")
    result = ReplaceFunc(result, "MATCH", "ПОИСКПОЗ")
    result = ReplaceFunc(result, "OFFSET", "СМЕЩ")
    result = ReplaceFunc(result, "ROWS", "ЧСТРОК")
    result = ReplaceFunc(result, "ROW", "СТРОКА")
    result = ReplaceFunc(result, "RTD", "ДРВ")
    result = ReplaceFunc(result, "TRANSPOSE", "ТРАНСП")
    result = ReplaceFunc(result, "VLOOKUP", "ВПР")
    result = ReplaceFunc(result, "XLOOKUP", "ПРОСМОТРX")
    result = ReplaceFunc(result, "XMATCH", "ПОИСКПОЗX")

    ' === СТАТИСТИЧЕСКИЕ ===
    result = ReplaceFunc(result, "AVEDEV", "СРОТКЛ")
    result = ReplaceFunc(result, "AVERAGEIFS", "СРЗНАЧЕСЛИМН")
    result = ReplaceFunc(result, "AVERAGEIF", "СРЗНАЧЕСЛИ")
    result = ReplaceFunc(result, "AVERAGEA", "СРЗНАЧА")
    result = ReplaceFunc(result, "AVERAGE", "СРЗНАЧ")
    result = ReplaceFunc(result, "BETA.DIST", "БЕТА.РАСП")
    result = ReplaceFunc(result, "BETA.INV", "БЕТА.ОБР")
    result = ReplaceFunc(result, "BINOM.DIST.RANGE", "БИНОМ.РАСП.ДИАП")
    result = ReplaceFunc(result, "BINOM.DIST", "БИНОМ.РАСП")
    result = ReplaceFunc(result, "BINOM.INV", "БИНОМ.ОБР")
    result = ReplaceFunc(result, "CHISQ.DIST.RT", "ХИ2.РАСП.ПХ")
    result = ReplaceFunc(result, "CHISQ.DIST", "ХИ2.РАСП")
    result = ReplaceFunc(result, "CHISQ.INV.RT", "ХИ2.ОБР.ПХ")
    result = ReplaceFunc(result, "CHISQ.INV", "ХИ2.ОБР")
    result = ReplaceFunc(result, "CHISQ.TEST", "ХИ2.ТЕСТ")
    result = ReplaceFunc(result, "CONFIDENCE.NORM", "ДОВЕРИТ.НОРМ")
    result = ReplaceFunc(result, "CONFIDENCE.T", "ДОВЕРИТ.СТЬЮДЕНТ")
    result = ReplaceFunc(result, "CORREL", "КОРРЕЛ")
    result = ReplaceFunc(result, "COUNTA", "СЧЁТЗ")
    result = ReplaceFunc(result, "COUNTBLANK", "СЧИТАТЬПУСТОТЫ")
    result = ReplaceFunc(result, "COUNTIFS", "СЧЁТЕСЛИМН")
    result = ReplaceFunc(result, "COUNTIF", "СЧЁТЕСЛИ")
    result = ReplaceFunc(result, "COUNT", "СЧЁТ")
    result = ReplaceFunc(result, "COVARIANCE.P", "КОВАРИАЦИЯ.Г")
    result = ReplaceFunc(result, "COVARIANCE.S", "КОВАРИАЦИЯ.В")
    result = ReplaceFunc(result, "DEVSQ", "КВАДРОТКЛ")
    result = ReplaceFunc(result, "EXPON.DIST", "ЭКСП.РАСП")
    result = ReplaceFunc(result, "F.DIST.RT", "F.РАСП.ПХ")
    result = ReplaceFunc(result, "F.DIST", "F.РАСП")
    result = ReplaceFunc(result, "F.INV.RT", "F.ОБР.ПХ")
    result = ReplaceFunc(result, "F.INV", "F.ОБР")
    result = ReplaceFunc(result, "FISHER", "ФИШЕР")
    result = ReplaceFunc(result, "FISHERINV", "ФИШЕРОБР")
    result = ReplaceFunc(result, "FORECAST.ETS.CONFINT", "ПРЕДСКАЗ.ETS.ДОВИНТЕРВАЛ")
    result = ReplaceFunc(result, "FORECAST.ETS.SEASONALITY", "ПРЕДСКАЗ.ETS.СЕЗОННОСТЬ")
    result = ReplaceFunc(result, "FORECAST.ETS.STAT", "ПРЕДСКАЗ.ETS.СТАТ")
    result = ReplaceFunc(result, "FORECAST.ETS", "ПРЕДСКАЗ.ETS")
    result = ReplaceFunc(result, "FORECAST.LINEAR", "ПРЕДСКАЗ")
    result = ReplaceFunc(result, "FORECAST", "ПРЕДСКАЗ")
    result = ReplaceFunc(result, "FREQUENCY", "ЧАСТОТА")
    result = ReplaceFunc(result, "F.TEST", "F.ТЕСТ")
    result = ReplaceFunc(result, "GAMMA.DIST", "ГАММА.РАСП")
    result = ReplaceFunc(result, "GAMMA.INV", "ГАММА.ОБР")
    result = ReplaceFunc(result, "GAMMALN.PRECISE", "ГАММАЛН.ТОЧН")
    result = ReplaceFunc(result, "GAMMALN", "ГАММАЛН")
    result = ReplaceFunc(result, "GAMMA", "ГАММА")
    result = ReplaceFunc(result, "GAUSS", "ГАУСС")
    result = ReplaceFunc(result, "GEOMEAN", "СРГЕОМ")
    result = ReplaceFunc(result, "GROWTH", "РОСТ")
    result = ReplaceFunc(result, "HARMEAN", "СРГАРМ")
    result = ReplaceFunc(result, "HYPGEOM.DIST", "ГИПЕРГЕОМ.РАСП")
    result = ReplaceFunc(result, "INTERCEPT", "ОТРЕЗОК")
    result = ReplaceFunc(result, "KURT", "ЭКСЦЕСС")
    result = ReplaceFunc(result, "LARGE", "НАИБОЛЬШИЙ")
    result = ReplaceFunc(result, "LINEST", "ЛИНЕЙН")
    result = ReplaceFunc(result, "LOGEST", "ЛГРФПРИБЛ")
    result = ReplaceFunc(result, "LOGNORM.DIST", "ЛОГНОРМ.РАСП")
    result = ReplaceFunc(result, "LOGNORM.INV", "ЛОГНОРМ.ОБР")
    result = ReplaceFunc(result, "MAXA", "МАКСА")
    result = ReplaceFunc(result, "MAXIFS", "МАКСЕСЛИМН")
    result = ReplaceFunc(result, "MAX", "МАКС")
    result = ReplaceFunc(result, "MEDIAN", "МЕДИАНА")
    result = ReplaceFunc(result, "MINA", "МИНА")
    result = ReplaceFunc(result, "MINIFS", "МИНЕСЛИМН")
    result = ReplaceFunc(result, "MIN", "МИН")
    result = ReplaceFunc(result, "MODE.MULT", "МОДА.НСК")
    result = ReplaceFunc(result, "MODE.SNGL", "МОДА.ОДН")
    result = ReplaceFunc(result, "MODE", "МОДА")
    result = ReplaceFunc(result, "NEGBINOM.DIST", "ОТРБИНОМ.РАСП")
    result = ReplaceFunc(result, "NORM.DIST", "НОРМ.РАСП")
    result = ReplaceFunc(result, "NORM.INV", "НОРМ.ОБР")
    result = ReplaceFunc(result, "NORM.S.DIST", "НОРМ.СТ.РАСП")
    result = ReplaceFunc(result, "NORM.S.INV", "НОРМ.СТ.ОБР")
    result = ReplaceFunc(result, "PEARSON", "ПИРСОН")
    result = ReplaceFunc(result, "PERCENTILE.EXC", "ПРОЦЕНТИЛЬ.ИСКЛ")
    result = ReplaceFunc(result, "PERCENTILE.INC", "ПРОЦЕНТИЛЬ.ВКЛ")
    result = ReplaceFunc(result, "PERCENTILE", "ПРОЦЕНТИЛЬ")
    result = ReplaceFunc(result, "PERCENTRANK.EXC", "ПРОЦЕНТРАНГ.ИСКЛ")
    result = ReplaceFunc(result, "PERCENTRANK.INC", "ПРОЦЕНТРАНГ.ВКЛ")
    result = ReplaceFunc(result, "PERCENTRANK", "ПРОЦЕНТРАНГ")
    result = ReplaceFunc(result, "PERMUT", "ПЕРЕСТ")
    result = ReplaceFunc(result, "PERMUTATIONA", "ПЕРЕСТА")
    result = ReplaceFunc(result, "PHI", "ФИ")
    result = ReplaceFunc(result, "POISSON.DIST", "ПУАССОН.РАСП")
    result = ReplaceFunc(result, "PROB", "ВЕРОЯТНОСТЬ")
    result = ReplaceFunc(result, "QUARTILE.EXC", "КВАРТИЛЬ.ИСКЛ")
    result = ReplaceFunc(result, "QUARTILE.INC", "КВАРТИЛЬ.ВКЛ")
    result = ReplaceFunc(result, "QUARTILE", "КВАРТИЛЬ")
    result = ReplaceFunc(result, "RANK.AVG", "РАНГ.СР")
    result = ReplaceFunc(result, "RANK.EQ", "РАНГ.РВ")
    result = ReplaceFunc(result, "RANK", "РАНГ")
    result = ReplaceFunc(result, "RSQ", "КВПИРСОН")
    result = ReplaceFunc(result, "SKEW.P", "СКОС.Г")
    result = ReplaceFunc(result, "SKEW", "СКОС")
    result = ReplaceFunc(result, "SLOPE", "НАКЛОН")
    result = ReplaceFunc(result, "SMALL", "НАИМЕНЬШИЙ")
    result = ReplaceFunc(result, "STANDARDIZE", "НОРМАЛИЗАЦИЯ")
    result = ReplaceFunc(result, "STDEV.P", "СТАНДОТКЛОН.Г")
    result = ReplaceFunc(result, "STDEV.S", "СТАНДОТКЛОН.В")
    result = ReplaceFunc(result, "STDEVA", "СТАНДОТКЛОНА")
    result = ReplaceFunc(result, "STDEVPA", "СТАНДОТКЛОНПА")
    result = ReplaceFunc(result, "STDEVP", "СТАНДОТКЛОНП")
    result = ReplaceFunc(result, "STDEV", "СТАНДОТКЛОН")
    result = ReplaceFunc(result, "STEYX", "СТОШYX")
    result = ReplaceFunc(result, "T.DIST.2T", "СТЬЮДЕНТ.РАСП.2Х")
    result = ReplaceFunc(result, "T.DIST.RT", "СТЬЮДЕНТ.РАСП.ПХ")
    result = ReplaceFunc(result, "T.DIST", "СТЬЮДЕНТ.РАСП")
    result = ReplaceFunc(result, "TREND", "ТЕНДЕНЦИЯ")
    result = ReplaceFunc(result, "TRIMMEAN", "УРЕЗСРЕДНЕЕ")
    result = ReplaceFunc(result, "T.INV.2T", "СТЬЮДЕНТ.ОБР.2Х")
    result = ReplaceFunc(result, "T.INV", "СТЬЮДЕНТ.ОБР")
    result = ReplaceFunc(result, "T.TEST", "СТЬЮДЕНТ.ТЕСТ")
    result = ReplaceFunc(result, "VAR.P", "ДИСП.Г")
    result = ReplaceFunc(result, "VAR.S", "ДИСП.В")
    result = ReplaceFunc(result, "VARA", "ДИСПА")
    result = ReplaceFunc(result, "VARPA", "ДИСПРА")
    result = ReplaceFunc(result, "VARP", "ДИСПР")
    result = ReplaceFunc(result, "VAR", "ДИСП")
    result = ReplaceFunc(result, "WEIBULL.DIST", "ВЕЙБУЛЛ.РАСП")
    result = ReplaceFunc(result, "Z.TEST", "Z.ТЕСТ")

    ' === ИНФОРМАЦИОННЫЕ ===
    result = ReplaceFunc(result, "CELL", "ЯЧЕЙКА")
    result = ReplaceFunc(result, "ERROR.TYPE", "ТИП.ОШИБКИ")
    result = ReplaceFunc(result, "INFO", "ИНФОРМ")
    result = ReplaceFunc(result, "ISBLANK", "ЕПУСТО")
    result = ReplaceFunc(result, "ISERR", "ЕОШ")
    result = ReplaceFunc(result, "ISERROR", "ЕОШИБКА")
    result = ReplaceFunc(result, "ISEVEN", "ЕЧЁТН")
    result = ReplaceFunc(result, "ISFORMULA", "ЕФОРМУЛА")
    result = ReplaceFunc(result, "ISLOGICAL", "ЕЛОГИЧ")
    result = ReplaceFunc(result, "ISNA", "ЕНД")
    result = ReplaceFunc(result, "ISNONTEXT", "ЕНЕТЕКСТ")
    result = ReplaceFunc(result, "ISNUMBER", "ЕЧИСЛО")
    result = ReplaceFunc(result, "ISODD", "ЕНЕЧЁТ")
    result = ReplaceFunc(result, "ISREF", "ЕССЫЛКА")
    result = ReplaceFunc(result, "ISTEXT", "ЕТЕКСТ")
    result = ReplaceFunc(result, "NA", "НД")
    result = ReplaceFunc(result, "SHEET", "ЛИСТ")
    result = ReplaceFunc(result, "SHEETS", "ЛИСТЫ")
    result = ReplaceFunc(result, "TYPE", "ТИП")

    ' === ФИНАНСОВЫЕ ===
    result = ReplaceFunc(result, "ACCRINT", "НАКОПДОХОД")
    result = ReplaceFunc(result, "ACCRINTM", "НАКОПДОХОДПОГАШ")
    result = ReplaceFunc(result, "AMORDEGRC", "АМОРУМ")
    result = ReplaceFunc(result, "AMORLINC", "АМОРУВ")
    result = ReplaceFunc(result, "COUPDAYBS", "ДНЕЙКУПОНДО")
    result = ReplaceFunc(result, "COUPDAYS", "ДНЕЙКУПОН")
    result = ReplaceFunc(result, "COUPDAYSNC", "ДНЕЙКУПОНПОСЛЕ")
    result = ReplaceFunc(result, "COUPNCD", "ДАТАКУПОНПОСЛЕ")
    result = ReplaceFunc(result, "COUPNUM", "ЧИСЛКУПОН")
    result = ReplaceFunc(result, "COUPPCD", "ДАТАКУПОНДО")
    result = ReplaceFunc(result, "CUMIPMT", "ОБЩПЛАТ")
    result = ReplaceFunc(result, "CUMPRINC", "ОБЩДОХОД")
    result = ReplaceFunc(result, "DB", "ФУО")
    result = ReplaceFunc(result, "DDB", "ДДОБ")
    result = ReplaceFunc(result, "DISC", "СКИДКА")
    result = ReplaceFunc(result, "DOLLARDE", "РУБЛЬ.ДЕС")
    result = ReplaceFunc(result, "DOLLARFR", "РУБЛЬ.ДРОБЬ")
    result = ReplaceFunc(result, "DURATION", "ДЛИТ")
    result = ReplaceFunc(result, "EFFECT", "ЭФФЕКТ")
    result = ReplaceFunc(result, "FV", "БС")
    result = ReplaceFunc(result, "FVSCHEDULE", "БЗРАСПИС")
    result = ReplaceFunc(result, "INTRATE", "ИНОРМА")
    result = ReplaceFunc(result, "IPMT", "ПРПЛТ")
    result = ReplaceFunc(result, "IRR", "ВСД")
    result = ReplaceFunc(result, "ISPMT", "ПРОЦПЛАТ")
    result = ReplaceFunc(result, "MDURATION", "МДЛИТ")
    result = ReplaceFunc(result, "MIRR", "МВСД")
    result = ReplaceFunc(result, "NOMINAL", "НОМИНАЛ")
    result = ReplaceFunc(result, "NPER", "КПЕР")
    result = ReplaceFunc(result, "NPV", "ЧПС")
    result = ReplaceFunc(result, "ODDFPRICE", "ЦЕНАПЕРВНЕРЕГ")
    result = ReplaceFunc(result, "ODDFYIELD", "ДОХОДПЕРВНЕРЕГ")
    result = ReplaceFunc(result, "ODDLPRICE", "ЦЕНАПОСЛНЕРЕГ")
    result = ReplaceFunc(result, "ODDLYIELD", "ДОХОДПОСЛНЕРЕГ")
    result = ReplaceFunc(result, "PDURATION", "ПДЛИТ")
    result = ReplaceFunc(result, "PMT", "ПЛТ")
    result = ReplaceFunc(result, "PPMT", "ОСПЛТ")
    result = ReplaceFunc(result, "PRICEDISC", "ЦЕНАСКИДКА")
    result = ReplaceFunc(result, "PRICEMAT", "ЦЕНАПОГАШ")
    result = ReplaceFunc(result, "PRICE", "ЦЕНА")
    result = ReplaceFunc(result, "PV", "ПС")
    result = ReplaceFunc(result, "RATE", "СТАВКА")
    result = ReplaceFunc(result, "RECEIVED", "ПОЛУЧЕНО")
    result = ReplaceFunc(result, "RRI", "ЭКВ.СТАВКА")
    result = ReplaceFunc(result, "SLN", "АПЛ")
    result = ReplaceFunc(result, "SYD", "АСЧ")
    result = ReplaceFunc(result, "TBILLEQ", "РАВНОКЧЕК")
    result = ReplaceFunc(result, "TBILLPRICE", "ЦЕНАКЧЕК")
    result = ReplaceFunc(result, "TBILLYIELD", "ДОХОДКЧЕК")
    result = ReplaceFunc(result, "VDB", "ПУО")
    result = ReplaceFunc(result, "XIRR", "ЧИСТВНДОХ")
    result = ReplaceFunc(result, "XNPV", "ЧИСТНЗ")
    result = ReplaceFunc(result, "YIELDDISC", "ДОХОДСКИДКА")
    result = ReplaceFunc(result, "YIELDMAT", "ДОХОДПОГАШ")
    result = ReplaceFunc(result, "YIELD", "ДОХОД")

    ' === ИНЖЕНЕРНЫЕ ===
    result = ReplaceFunc(result, "BESSELI", "БЕССЕЛЬ.I")
    result = ReplaceFunc(result, "BESSELJ", "БЕССЕЛЬ.J")
    result = ReplaceFunc(result, "BESSELK", "БЕССЕЛЬ.K")
    result = ReplaceFunc(result, "BESSELY", "БЕССЕЛЬ.Y")
    result = ReplaceFunc(result, "BIN2DEC", "ДВ.В.ДЕС")
    result = ReplaceFunc(result, "BIN2HEX", "ДВ.В.ШЕСТН")
    result = ReplaceFunc(result, "BIN2OCT", "ДВ.В.ВОСЬМ")
    result = ReplaceFunc(result, "BITAND", "БИТ.И")
    result = ReplaceFunc(result, "BITLSHIFT", "БИТ.СДВИГЛ")
    result = ReplaceFunc(result, "BITOR", "БИТ.ИЛИ")
    result = ReplaceFunc(result, "BITRSHIFT", "БИТ.СДВИГП")
    result = ReplaceFunc(result, "BITXOR", "БИТ.ИСКЛИЛИ")
    result = ReplaceFunc(result, "COMPLEX", "КОМПЛЕКСН")
    result = ReplaceFunc(result, "CONVERT", "ПРЕОБР")
    result = ReplaceFunc(result, "DEC2BIN", "ДЕС.В.ДВ")
    result = ReplaceFunc(result, "DEC2HEX", "ДЕС.В.ШЕСТН")
    result = ReplaceFunc(result, "DEC2OCT", "ДЕС.В.ВОСЬМ")
    result = ReplaceFunc(result, "DELTA", "ДЕЛЬТА")
    result = ReplaceFunc(result, "ERF.PRECISE", "ФОШ.ТОЧН")
    result = ReplaceFunc(result, "ERFC.PRECISE", "ДФОШ.ТОЧН")
    result = ReplaceFunc(result, "ERFC", "ДФОШ")
    result = ReplaceFunc(result, "ERF", "ФОШ")
    result = ReplaceFunc(result, "GESTEP", "ПОРОГ")
    result = ReplaceFunc(result, "HEX2BIN", "ШЕСТН.В.ДВ")
    result = ReplaceFunc(result, "HEX2DEC", "ШЕСТН.В.ДЕС")
    result = ReplaceFunc(result, "HEX2OCT", "ШЕСТН.В.ВОСЬМ")
    result = ReplaceFunc(result, "IMABS", "МНИМ.ABS")
    result = ReplaceFunc(result, "IMAGINARY", "МНИМ.ЧАСТЬ")
    result = ReplaceFunc(result, "IMARGUMENT", "МНИМ.АРГУМЕНТ")
    result = ReplaceFunc(result, "IMCONJUGATE", "МНИМ.СОПРЯЖ")
    result = ReplaceFunc(result, "IMCOS", "МНИМ.COS")
    result = ReplaceFunc(result, "IMCOSH", "МНИМ.COSH")
    result = ReplaceFunc(result, "IMCOT", "МНИМ.COT")
    result = ReplaceFunc(result, "IMCSC", "МНИМ.CSC")
    result = ReplaceFunc(result, "IMCSCH", "МНИМ.CSCH")
    result = ReplaceFunc(result, "IMDIV", "МНИМ.ДЕЛ")
    result = ReplaceFunc(result, "IMEXP", "МНИМ.EXP")
    result = ReplaceFunc(result, "IMLN", "МНИМ.LN")
    result = ReplaceFunc(result, "IMLOG10", "МНИМ.LOG10")
    result = ReplaceFunc(result, "IMLOG2", "МНИМ.LOG2")
    result = ReplaceFunc(result, "IMPOWER", "МНИМ.СТЕПЕНЬ")
    result = ReplaceFunc(result, "IMPRODUCT", "МНИМ.ПРОИЗВЕД")
    result = ReplaceFunc(result, "IMREAL", "МНИМ.ВЕЩ")
    result = ReplaceFunc(result, "IMSEC", "МНИМ.SEC")
    result = ReplaceFunc(result, "IMSECH", "МНИМ.SECH")
    result = ReplaceFunc(result, "IMSIN", "МНИМ.SIN")
    result = ReplaceFunc(result, "IMSINH", "МНИМ.SINH")
    result = ReplaceFunc(result, "IMSQRT", "МНИМ.КОРЕНЬ")
    result = ReplaceFunc(result, "IMSUB", "МНИМ.РАЗН")
    result = ReplaceFunc(result, "IMSUM", "МНИМ.СУММ")
    result = ReplaceFunc(result, "IMTAN", "МНИМ.TAN")
    result = ReplaceFunc(result, "OCT2BIN", "ВОСЬМ.В.ДВ")
    result = ReplaceFunc(result, "OCT2DEC", "ВОСЬМ.В.ДЕС")
    result = ReplaceFunc(result, "OCT2HEX", "ВОСЬМ.В.ШЕСТН")

    ' === БАЗЫ ДАННЫХ ===
    result = ReplaceFunc(result, "DAVERAGE", "ДСРЗНАЧ")
    result = ReplaceFunc(result, "DCOUNT", "БСЧЁТ")
    result = ReplaceFunc(result, "DCOUNTA", "БСЧЁТА")
    result = ReplaceFunc(result, "DGET", "БИЗВЛЕЧЬ")
    result = ReplaceFunc(result, "DMAX", "ДМАКС")
    result = ReplaceFunc(result, "DMIN", "ДМИН")
    result = ReplaceFunc(result, "DPRODUCT", "БДПРОИЗВЕД")
    result = ReplaceFunc(result, "DSTDEVP", "ДСТАНДОТКЛОНП")
    result = ReplaceFunc(result, "DSTDEV", "ДСТАНДОТКЛ")
    result = ReplaceFunc(result, "DSUM", "БДСУММ")
    result = ReplaceFunc(result, "DVARP", "БДДИСПП")
    result = ReplaceFunc(result, "DVAR", "БДДИСП")

    ' === ВЕБ ===
    result = ReplaceFunc(result, "ENCODEURL", "КОДИР.URL")
    result = ReplaceFunc(result, "FILTERXML", "ФИЛЬТР.XML")
    result = ReplaceFunc(result, "WEBSERVICE", "ВЕБСЛУЖБА")

    ' === ДИНАМИЧЕСКИЕ МАССИВЫ (Excel 365) ===
    result = ReplaceFunc(result, "FILTER", "ФИЛЬТР")
    result = ReplaceFunc(result, "RANDARRAY", "СЛУЧМАССИВ")
    result = ReplaceFunc(result, "SEQUENCE", "ПОСЛЕДОВ")
    result = ReplaceFunc(result, "SORTBY", "СОРТПО")
    result = ReplaceFunc(result, "SORT", "СОРТ")
    result = ReplaceFunc(result, "UNIQUE", "УНИК")

    ' === ПРОЧИЕ ===
    result = ReplaceFunc(result, "EUROCONVERT", "ЕВРО")
    result = ReplaceFunc(result, "N", "Ч")
    result = ReplaceFunc(result, "T", "Т")

    ' Заменяем константы TRUE/FALSE
    result = ReplaceConstant(result, "TRUE", "ИСТИНА")
    result = ReplaceConstant(result, "FALSE", "ЛОЖЬ")

    ' Заменяем разделитель аргументов (запятая -> точка с запятой для русской локали)
    result = Replace(result, ",", Application.International(xlListSeparator))

    LocalizeFormula = result
End Function


'----------------------------------------
' Преобразование буквы столбца или числа в номер колонки относительно диапазона
' colRef - может быть числом (1, 2, 3) или буквой (A, B, C)
' rng - диапазон, относительно которого вычисляется номер
' Возвращает номер колонки в диапазоне (1-based)
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: colRef, rng, ws
' Выход: значение функции
' ---
Private Function GetColumnNumber(colRef As String, rng As Range, ws As Worksheet) As Long
    Dim col As String
    col = Trim(colRef)

    If IsNumeric(col) Then
        ' Уже число - возвращаем как есть
        GetColumnNumber = CLng(col)
    Else
        ' Буква столбца - вычисляем номер относительно начала диапазона
        On Error Resume Next
        Dim absColNum As Long
        absColNum = ws.columns(col).Column ' Абсолютный номер столбца (A=1, B=2...)

        If absColNum > 0 Then
            ' Вычисляем относительный номер в диапазоне
            GetColumnNumber = absColNum - rng.Column + 1
            If GetColumnNumber < 1 Then GetColumnNumber = 1
            If GetColumnNumber > rng.columns.Count Then GetColumnNumber = rng.columns.Count
        Else
            GetColumnNumber = 1 ' По умолчанию первая колонка
        End If
        On Error GoTo 0
    End If
End Function


'----------------------------------------
' Замена имени функции (с учётом что это функция, а не часть текста)
'----------------------------------------
' ---
' Что делает: Вспомогательная функция модуля.
' Вход: formula, engName, rusName
' Выход: значение функции
' ---
Private Function ReplaceFunc(formula As String, engName As String, rusName As String) As String
    Dim result As String
    Dim pos As Long
    Dim before As String

    result = formula
    pos = InStr(1, UCase(result), UCase(engName) & "(")

    Do While pos > 0
        ' Проверяем, что перед именем функции нет буквы (чтобы не заменить часть другой функции)
        If pos = 1 Then
            result = rusName & Mid(result, pos + Len(engName))
        Else
            before = Mid(result, pos - 1, 1)
            If Not (before >= "A" And before <= "Z") And Not (before >= "a" And before <= "z") And Not (before >= "А" And before <= "я") Then
                result = Left(result, pos - 1) & rusName & Mid(result, pos + Len(engName))
            End If
        End If
        pos = InStr(pos + Len(rusName), UCase(result), UCase(engName) & "(")
    Loop

    ReplaceFunc = result
End Function


'----------------------------------------
' Замена константы (TRUE, FALSE) - без требования скобки после имени
'----------------------------------------
' ---
' Что делает: Вспомогательная функция модуля.
' Вход: formula, engName, rusName
' Выход: значение функции
' ---
Private Function ReplaceConstant(formula As String, engName As String, rusName As String) As String
    Dim result As String
    Dim pos As Long
    Dim before As String
    Dim after As String
    Dim isWordBoundary As Boolean

    result = formula
    pos = InStr(1, UCase(result), UCase(engName))

    Do While pos > 0
        isWordBoundary = True

        ' Проверяем символ перед
        If pos > 1 Then
            before = Mid(result, pos - 1, 1)
            If (before >= "A" And before <= "Z") Or (before >= "a" And before <= "z") Or (before >= "А" And before <= "я") Or (before >= "0" And before <= "9") Then
                isWordBoundary = False
            End If
        End If

        ' Проверяем символ после
        If pos + Len(engName) <= Len(result) Then
            after = Mid(result, pos + Len(engName), 1)
            If (after >= "A" And after <= "Z") Or (after >= "a" And after <= "z") Or (after >= "А" And after <= "я") Or (after >= "0" And after <= "9") Then
                isWordBoundary = False
            End If
        End If

        If isWordBoundary Then
            result = Left(result, pos - 1) & rusName & Mid(result, pos + Len(engName))
            pos = InStr(pos + Len(rusName), UCase(result), UCase(engName))
        Else
            pos = InStr(pos + 1, UCase(result), UCase(engName))
        End If
    Loop

    ReplaceConstant = result
End Function


'----------------------------------------
' Получение типа графика
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: chartTypeName
' Выход: значение функции
' ---
Private Function GetChartType(chartTypeName As String) As Long
    Select Case UCase(Trim(chartTypeName))
        Case "LINE": GetChartType = 4 ' xlLine
        Case "BAR": GetChartType = 57 ' xlBarClustered
        Case "COLUMN": GetChartType = 51 ' xlColumnClustered
        Case "PIE": GetChartType = 5 ' xlPie
        Case "AREA": GetChartType = 1 ' xlArea
        Case "SCATTER", "XY": GetChartType = -4169 ' xlXYScatter
        Case "DOUGHNUT": GetChartType = -4120 ' xlDoughnut
        Case "RADAR": GetChartType = -4151 ' xlRadar
        Case "SURFACE": GetChartType = 83 ' xlSurface
        Case "BUBBLE": GetChartType = 15 ' xlBubble
        Case "STOCK": GetChartType = 88 ' xlStockHLC
        Case "CYLINDER": GetChartType = 95 ' xlCylinderCol
        Case "CONE": GetChartType = 99 ' xlConeCol
        Case "PYRAMID": GetChartType = 103 ' xlPyramidCol
        Case "LINE_MARKERS": GetChartType = 65 ' xlLineMarkers
        Case "AREA_STACKED": GetChartType = 76 ' xlAreaStacked
        Case "BAR_STACKED": GetChartType = 58 ' xlBarStacked
        Case "COLUMN_STACKED": GetChartType = 52 ' xlColumnStacked
        Case Else: GetChartType = 4 ' xlLine по умолчанию
    End Select
End Function


'----------------------------------------
' Поиск сводной таблицы по имени во всех листах
'----------------------------------------
' ---
' Что делает: Вспомогательная функция модуля.
' Вход: pivotName
' Выход: значение функции
' ---
Private Function FindPivotTable(pivotName As String) As pivotTable
    Dim wsSearch As Worksheet
    Dim ptSearch As pivotTable
    Dim searchName As String

    searchName = Trim(pivotName)
    Set FindPivotTable = Nothing

    On Error Resume Next
    ' Ищем во всех листах книги
    For Each wsSearch In ActiveWorkbook.Worksheets
        For Each ptSearch In wsSearch.PivotTables
            If ptSearch.Name = searchName Then
                Set FindPivotTable = ptSearch
                Exit Function
            End If
        Next ptSearch
    Next wsSearch
    On Error GoTo 0
End Function


'----------------------------------------
' Получение индекса графика
' Поддерживает: 0, LAST, NEW = последний график; 1,2,3... = конкретный индекс
'----------------------------------------
' ---
' Что делает: Читает данные из Excel/настроек и возвращает результат.
' Вход: ws, indexStr
' Выход: значение функции
' ---
Private Function GetChartIndex(ws As Worksheet, indexStr As String) As Long
    Dim idx As Long
    Dim s As String

    s = UCase(Trim(indexStr))

    ' Если 0, LAST или NEW - возвращаем последний график
    If s = "0" Or s = "LAST" Or s = "NEW" Or s = "" Then
        GetChartIndex = ws.ChartObjects.Count
        Exit Function
    End If

    ' Пробуем преобразовать в число
    On Error Resume Next
    idx = CLng(s)
    On Error GoTo 0

    If idx > 0 Then
        GetChartIndex = idx
    Else
        ' По умолчанию - последний
        GetChartIndex = ws.ChartObjects.Count
    End If
End Function


'----------------------------------------
' Выполнение одной команды (ПОЛНАЯ ВЕРСИЯ)
'----------------------------------------
' ---
' Что делает: Выполняет подготовленные команды в Excel.
' Вход: cmd
' Выход: значение функции
' ---
Private Function ExecuteSingleCommand(cmd As String) As Boolean
    On Error GoTo ErrorHandler

    Dim parts() As String
    Dim action As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim i As Long

    parts = Split(cmd, "|")
    If UBound(parts) < 0 Then
        ExecuteSingleCommand = False
        Exit Function
    End If

    action = UCase(Trim(parts(0)))
    Set ws = ActiveSheet

    ' Отладка: показываем команду и количество частей
    Debug.Print "CMD: " & action & " | Parts: " & (UBound(parts) + 1) & " | Full: " & cmd

    Select Case action

        ' ========== РАБОТА С ЯЧЕЙКАМИ ==========

        Case "SET_VALUE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).value = parts(2)
                ExecuteSingleCommand = True
            End If

        Case "SET_FORMULA"
            If UBound(parts) >= 2 Then
                Dim localFormula As String
                localFormula = LocalizeFormula(parts(2))
                Debug.Print "SET_FORMULA: Original=" & parts(2)
                Debug.Print "SET_FORMULA: Localized=" & localFormula
                ' Используем FormulaLocal для русских формул
                ws.Range(parts(1)).FormulaLocal = localFormula
                ExecuteSingleCommand = True
            End If

        Case "FILL_DOWN"
            If UBound(parts) >= 2 Then
                Dim srcCell As Range, destRng As Range
                Set srcCell = ws.Range(parts(1))
                Set destRng = ws.Range(parts(1) & ":" & parts(2))
                srcCell.Copy
                destRng.PasteSpecial xlPasteFormulas
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If

        Case "FILL_RIGHT"
            If UBound(parts) >= 2 Then
                Dim srcCellR As Range, destRngR As Range
                Set srcCellR = ws.Range(parts(1))
                Set destRngR = ws.Range(parts(1) & ":" & parts(2))
                srcCellR.Copy
                destRngR.PasteSpecial xlPasteFormulas
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If

        Case "FILL_SERIES"
            If UBound(parts) >= 2 Then
                Dim stepVal As Double
                stepVal = 1
                If UBound(parts) >= 2 Then stepVal = CDbl(parts(2))
                ws.Range(parts(1)).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Step:=stepVal
                ExecuteSingleCommand = True
            End If

        Case "CLEAR_CONTENTS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).ClearContents
                ExecuteSingleCommand = True
            End If

        Case "CLEAR_FORMAT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).ClearFormats
                ExecuteSingleCommand = True
            End If

        Case "CLEAR_ALL"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Clear
                ExecuteSingleCommand = True
            End If

        Case "COPY"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy Destination:=ws.Range(parts(2))
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If

        Case "CUT"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Cut Destination:=ws.Range(parts(2))
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If

        Case "PASTE_VALUES"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy
                ws.Range(parts(2)).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If

        Case "TRANSPOSE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy
                ws.Range(parts(2)).PasteSpecial Paste:=xlPasteAll, Transpose:=True
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If

        ' ========== ФОРМАТИРОВАНИЕ ==========

        Case "BOLD"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Bold = True
                ExecuteSingleCommand = True
            End If

        Case "ITALIC"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Italic = True
                ExecuteSingleCommand = True
            End If

        Case "UNDERLINE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Underline = xlUnderlineStyleSingle
                ExecuteSingleCommand = True
            End If

        Case "STRIKETHROUGH"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Strikethrough = True
                ExecuteSingleCommand = True
            End If

        Case "FONT_NAME"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Name = parts(2)
                ExecuteSingleCommand = True
            End If

        Case "FONT_SIZE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Size = CLng(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "FONT_COLOR"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "FILL_COLOR"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "BORDER"
            If UBound(parts) >= 2 Then
                Dim borderStyle As String
                borderStyle = UCase(Trim(parts(2)))
                With ws.Range(parts(1))
                    Select Case borderStyle
                        Case "ALL"
                            .Borders.LineStyle = xlContinuous
                        Case "TOP"
                            .Borders(xlEdgeTop).LineStyle = xlContinuous
                        Case "BOTTOM"
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        Case "LEFT"
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        Case "RIGHT"
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                        Case "NONE"
                            .Borders.LineStyle = xlNone
                    End Select
                End With
                ExecuteSingleCommand = True
            End If

        Case "BORDER_THICK"
            If UBound(parts) >= 1 Then
                With ws.Range(parts(1)).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                ExecuteSingleCommand = True
            End If

        Case "ALIGN_H"
            If UBound(parts) >= 2 Then
                Dim hAlign As Long
                Select Case UCase(Trim(parts(2)))
                    Case "LEFT": hAlign = xlLeft
                    Case "CENTER": hAlign = xlCenter
                    Case "RIGHT": hAlign = xlRight
                    Case "JUSTIFY": hAlign = xlJustify
                    Case Else: hAlign = xlGeneral
                End Select
                ws.Range(parts(1)).HorizontalAlignment = hAlign
                ExecuteSingleCommand = True
            End If

        Case "ALIGN_V"
            If UBound(parts) >= 2 Then
                Dim vAlign As Long
                Select Case UCase(Trim(parts(2)))
                    Case "TOP": vAlign = xlTop
                    Case "CENTER": vAlign = xlCenter
                    Case "BOTTOM": vAlign = xlBottom
                    Case Else: vAlign = xlCenter
                End Select
                ws.Range(parts(1)).VerticalAlignment = vAlign
                ExecuteSingleCommand = True
            End If

        Case "WRAP_TEXT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).WrapText = True
                ExecuteSingleCommand = True
            End If

        Case "MERGE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Merge
                ExecuteSingleCommand = True
            End If

        Case "UNMERGE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).UnMerge
                ExecuteSingleCommand = True
            End If

        Case "FORMAT_NUMBER"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).NumberFormat = parts(2)
                ExecuteSingleCommand = True
            End If

        Case "FORMAT_DATE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).NumberFormat = parts(2)
                ExecuteSingleCommand = True
            End If

        Case "FORMAT_PERCENT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).NumberFormat = "0.00%"
                ExecuteSingleCommand = True
            End If

        Case "FORMAT_CURRENCY"
            If UBound(parts) >= 1 Then
                Dim currSymbol As String
                currSymbol = "?"
                If UBound(parts) >= 2 Then currSymbol = parts(2)
                ws.Range(parts(1)).NumberFormat = "#,##0.00 " & currSymbol
                ExecuteSingleCommand = True
            End If

        Case "AUTOFIT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).columns.AutoFit
                ExecuteSingleCommand = True
            End If

        Case "AUTOFIT_ROWS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Rows.AutoFit
                ExecuteSingleCommand = True
            End If

        Case "COLUMN_WIDTH"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1)).ColumnWidth = CDbl(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "ROW_HEIGHT"
            If UBound(parts) >= 2 Then
                ws.Rows(CLng(parts(1))).RowHeight = CDbl(parts(2))
                ExecuteSingleCommand = True
            End If

        ' ========== СТРОКИ И СТОЛБЦЫ ==========

        Case "INSERT_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Insert
                ExecuteSingleCommand = True
            End If

        Case "INSERT_ROWS"
            If UBound(parts) >= 2 Then
                Dim rowNum As Long, rowCount As Long
                rowNum = CLng(parts(1))
                rowCount = CLng(parts(2))
                ws.Rows(rowNum & ":" & (rowNum + rowCount - 1)).Insert
                ExecuteSingleCommand = True
            End If

        Case "INSERT_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Insert
                ExecuteSingleCommand = True
            End If

        Case "INSERT_COLUMNS"
            If UBound(parts) >= 2 Then
                Dim colNum As Long, colCount As Long
                colCount = CLng(parts(2))
                For i = 1 To colCount
                    ws.columns(parts(1)).Insert
                Next i
                ExecuteSingleCommand = True
            End If

        Case "DELETE_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Delete
                ExecuteSingleCommand = True
            End If

        Case "DELETE_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Delete
                ExecuteSingleCommand = True
            End If

        Case "DELETE_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Delete
                ExecuteSingleCommand = True
            End If

        Case "DELETE_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Delete
                ExecuteSingleCommand = True
            End If

        Case "HIDE_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Hidden = True
                ExecuteSingleCommand = True
            End If

        Case "HIDE_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Hidden = True
                ExecuteSingleCommand = True
            End If

        Case "SHOW_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Hidden = False
                ExecuteSingleCommand = True
            End If

        Case "SHOW_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Hidden = False
                ExecuteSingleCommand = True
            End If

        Case "HIDE_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Hidden = True
                ExecuteSingleCommand = True
            End If

        Case "SHOW_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Hidden = False
                ExecuteSingleCommand = True
            End If

        Case "GROUP_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Group
                ExecuteSingleCommand = True
            End If

        Case "UNGROUP_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Ungroup
                ExecuteSingleCommand = True
            End If

        Case "GROUP_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Group
                ExecuteSingleCommand = True
            End If

        Case "UNGROUP_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Ungroup
                ExecuteSingleCommand = True
            End If

        ' ========== СОРТИРОВКА И ФИЛЬТРАЦИЯ ==========

        Case "SORT"
            If UBound(parts) >= 3 Then
                Dim SortRange As Range, sortCol As Long, sortOrder As Long
                Dim sortColStr As String, sortKeyRange As Range
                Set SortRange = ws.Range(parts(1))
                sortColStr = Trim(parts(2))

                ' Определяем колонку для сортировки
                If IsNumeric(sortColStr) Then
                    ' Число - номер колонки в диапазоне (1, 2, 3...)
                    sortCol = CLng(sortColStr)
                    Set sortKeyRange = SortRange.columns(sortCol)
                Else
                    ' Буква столбца (A, B, C...) - используем пересечение с диапазоном
                    Set sortKeyRange = Intersect(SortRange, ws.columns(sortColStr))
                    If sortKeyRange Is Nothing Then
                        ' Если буква вне диапазона, берём первую колонку
                        Set sortKeyRange = SortRange.columns(1)
                    End If
                End If

                sortOrder = IIf(UCase(parts(3)) = "ASC", xlAscending, xlDescending)
                SortRange.Sort Key1:=sortKeyRange, Order1:=sortOrder, Header:=xlGuess
                ExecuteSingleCommand = True
            End If

        Case "SORT_MULTI"
            If UBound(parts) >= 5 Then
                Dim sRng As Range
                Dim sCol1 As Long, sCol2 As Long
                Set sRng = ws.Range(parts(1))
                Dim o1 As Long, o2 As Long
                sCol1 = GetColumnNumber(parts(2), sRng, ws)
                sCol2 = GetColumnNumber(parts(4), sRng, ws)
                o1 = IIf(UCase(parts(3)) = "ASC", xlAscending, xlDescending)
                o2 = IIf(UCase(parts(5)) = "ASC", xlAscending, xlDescending)
                sRng.Sort Key1:=sRng.columns(sCol1), Order1:=o1, _
                          Key2:=sRng.columns(sCol2), Order2:=o2, Header:=xlGuess
                ExecuteSingleCommand = True
            End If

        Case "AUTOFILTER"
            If UBound(parts) >= 1 Then
                If ws.AutoFilterMode Then ws.AutoFilterMode = False
                ws.Range(parts(1)).AutoFilter
                ExecuteSingleCommand = True
            End If

        Case "FILTER"
            If UBound(parts) >= 3 Then
                Dim fRng As Range
                Dim fCol As Long
                Set fRng = ws.Range(parts(1))
                fCol = GetColumnNumber(parts(2), fRng, ws)
                If Not ws.AutoFilterMode Then fRng.AutoFilter
                fRng.AutoFilter Field:=fCol, Criteria1:=parts(3)
                ExecuteSingleCommand = True
            End If

        Case "FILTER_TOP"
            If UBound(parts) >= 3 Then
                Dim ftRng As Range
                Dim ftCol As Long
                Set ftRng = ws.Range(parts(1))
                ftCol = GetColumnNumber(parts(2), ftRng, ws)
                If Not ws.AutoFilterMode Then ftRng.AutoFilter
                ftRng.AutoFilter Field:=ftCol, Criteria1:=CLng(parts(3)), Operator:=xlTop10Items
                ExecuteSingleCommand = True
            End If

        Case "CLEAR_FILTER"
            If UBound(parts) >= 1 Then
                If ws.AutoFilterMode Then
                    ws.Range(parts(1)).AutoFilter
                    ws.Range(parts(1)).AutoFilter
                End If
                ExecuteSingleCommand = True
            End If

        Case "REMOVE_AUTOFILTER"
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            ExecuteSingleCommand = True

        Case "REMOVE_DUPLICATES"
            If UBound(parts) >= 2 Then
                Dim dupRng As Range
                Dim colsArr() As Long
                Dim colsList() As String
                Set dupRng = ws.Range(parts(1))
                colsList = Split(parts(2), ",")
                ReDim colsArr(UBound(colsList))
                For i = 0 To UBound(colsList)
                    colsArr(i) = CLng(Trim(colsList(i)))
                Next i
                dupRng.RemoveDuplicates columns:=colsArr, Header:=xlYes
                ExecuteSingleCommand = True
            End If

        Case "FIND_REPLACE"
            If UBound(parts) >= 2 Then
                Dim replaceWith As String
                replaceWith = ""
                If UBound(parts) >= 2 Then replaceWith = parts(2)
                ws.UsedRange.Replace What:=parts(1), Replacement:=replaceWith, LookAt:=xlPart
                ExecuteSingleCommand = True
            End If

        Case "FIND_REPLACE_RANGE"
            If UBound(parts) >= 3 Then
                ws.Range(parts(1)).Replace What:=parts(2), Replacement:=parts(3), LookAt:=xlPart
                ExecuteSingleCommand = True
            End If

        ' ========== ГРАФИКИ ==========

        Case "CREATE_CHART"
            ' CREATE_CHART|диапазон|тип|название
            ' Поддерживает несмежные диапазоны: A2:A5,B2:B5
            If UBound(parts) >= 2 Then
                Dim chartObj As ChartObject
                Dim chartType As Long
                Dim dataRange As Range
                Dim chartLeft As Double, chartTop As Double
                Dim rangeStr As String
                Dim rangeParts() As String
                Dim rngPart As Variant

                rangeStr = Trim(parts(1))

                ' Проверяем на несмежные диапазоны
                On Error Resume Next
                If InStr(rangeStr, ",") > 0 Then
                    ' Несмежные диапазоны
                    rangeParts = Split(rangeStr, ",")
                    Set dataRange = ws.Range(Trim(rangeParts(0)))
                    For i = 1 To UBound(rangeParts)
                        Set dataRange = Union(dataRange, ws.Range(Trim(rangeParts(i))))
                    Next i
                Else
                    Set dataRange = ws.Range(rangeStr)
                End If
                On Error GoTo ErrorHandler

                If dataRange Is Nothing Then
                    ExecuteSingleCommand = False
                    Exit Function
                End If

                chartType = GetChartType(parts(2))

                ' Позиционируем график справа от данных
                chartLeft = dataRange.Areas(1).Cells(1, dataRange.Areas(1).columns.Count).Offset(0, 2).Left
                chartTop = dataRange.Areas(1).Top

                Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=400, Height:=250)
                chartObj.Chart.SetSourceData Source:=dataRange
                chartObj.Chart.chartType = chartType

                If UBound(parts) >= 3 Then
                    If Len(Trim(parts(3))) > 0 Then
                        chartObj.Chart.HasTitle = True
                        chartObj.Chart.ChartTitle.text = parts(3)
                    End If
                End If

                ExecuteSingleCommand = True
            End If

        Case "CREATE_CHART_POS", "CREATE_CHART_AT"
            ' CREATE_CHART_POS|диапазон|тип|название|ячейка_или_лево|верх|ширина|высота
            ' CREATE_CHART_AT|диапазон|тип|название|ячейка - упрощённый вариант
            ' Поддерживает несмежные диапазоны: A2:A5,B2:B5
            If UBound(parts) >= 3 Then
                Dim chartObj2 As ChartObject
                Dim chartType2 As Long
                Dim dataRange2 As Range
                Dim posLeft As Double, posTop As Double
                Dim chWidth As Double, chHeight As Double
                Dim rangeStr2 As String
                Dim rangeParts2() As String

                rangeStr2 = Trim(parts(1))

                ' Проверяем на несмежные диапазоны
                On Error Resume Next
                If InStr(rangeStr2, ",") > 0 Then
                    rangeParts2 = Split(rangeStr2, ",")
                    Set dataRange2 = ws.Range(Trim(rangeParts2(0)))
                    For i = 1 To UBound(rangeParts2)
                        Set dataRange2 = Union(dataRange2, ws.Range(Trim(rangeParts2(i))))
                    Next i
                Else
                    Set dataRange2 = ws.Range(rangeStr2)
                End If
                On Error GoTo ErrorHandler

                If dataRange2 Is Nothing Then
                    ExecuteSingleCommand = False
                    Exit Function
                End If

                chartType2 = GetChartType(parts(2))

                ' Значения по умолчанию
                chWidth = 400
                chHeight = 250
                posLeft = 300
                posTop = 50

                ' Определяем позицию
                If UBound(parts) >= 4 Then
                    ' Проверяем, это адрес ячейки или число
                    On Error Resume Next
                    Dim posCell As Range
                    Set posCell = ws.Range(parts(4))
                    If Not posCell Is Nothing Then
                        ' Это адрес ячейки
                        posLeft = posCell.Left
                        posTop = posCell.Top
                    Else
                        ' Это число
                        posLeft = CDbl(parts(4))
                    End If
                    On Error GoTo ErrorHandler
                End If

                If UBound(parts) >= 5 Then
                    On Error Resume Next
                    posTop = CDbl(parts(5))
                    On Error GoTo ErrorHandler
                End If

                If UBound(parts) >= 6 Then
                    On Error Resume Next
                    chWidth = CDbl(parts(6))
                    On Error GoTo ErrorHandler
                End If

                If UBound(parts) >= 7 Then
                    On Error Resume Next
                    chHeight = CDbl(parts(7))
                    On Error GoTo ErrorHandler
                End If

                Set chartObj2 = ws.ChartObjects.Add(Left:=posLeft, Top:=posTop, Width:=chWidth, Height:=chHeight)
                chartObj2.Chart.SetSourceData Source:=dataRange2
                chartObj2.Chart.chartType = chartType2

                If Len(Trim(parts(3))) > 0 Then
                    chartObj2.Chart.HasTitle = True
                    chartObj2.Chart.ChartTitle.text = parts(3)
                End If

                ExecuteSingleCommand = True
            End If

        Case "CHART_TITLE"
            ' CHART_TITLE|индекс|текст (индекс: 1 = первый график, 0 или LAST = последний)
            If UBound(parts) >= 2 Then
                Dim chIdx As Long
                chIdx = GetChartIndex(ws, parts(1))
                If chIdx > 0 And chIdx <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx).Chart.HasTitle = True
                    ws.ChartObjects(chIdx).Chart.ChartTitle.text = parts(2)
                End If
                ExecuteSingleCommand = True
            End If

        Case "CHART_LEGEND"
            If UBound(parts) >= 2 Then
                Dim chIdx2 As Long
                chIdx2 = GetChartIndex(ws, parts(1))
                If chIdx2 > 0 And chIdx2 <= ws.ChartObjects.Count Then
                    Dim legPos As Long
                    Select Case UCase(Trim(parts(2)))
                        Case "TOP": legPos = xlLegendPositionTop
                        Case "BOTTOM": legPos = xlLegendPositionBottom
                        Case "LEFT": legPos = xlLegendPositionLeft
                        Case "RIGHT": legPos = xlLegendPositionRight
                        Case "NONE"
                            ws.ChartObjects(chIdx2).Chart.HasLegend = False
                            ExecuteSingleCommand = True
                            Exit Function
                        Case Else: legPos = xlLegendPositionBottom
                    End Select
                    ws.ChartObjects(chIdx2).Chart.HasLegend = True
                    ws.ChartObjects(chIdx2).Chart.Legend.Position = legPos
                End If
                ExecuteSingleCommand = True
            End If

        Case "CHART_AXIS_TITLE"
            If UBound(parts) >= 3 Then
                Dim chIdx3 As Long
                chIdx3 = GetChartIndex(ws, parts(1))
                If chIdx3 > 0 And chIdx3 <= ws.ChartObjects.Count Then
                    On Error Resume Next
                    Dim ax As Object
                    If UCase(Trim(parts(2))) = "X" Then
                        Set ax = ws.ChartObjects(chIdx3).Chart.Axes(xlCategory)
                    Else
                        Set ax = ws.ChartObjects(chIdx3).Chart.Axes(xlValue)
                    End If
                    If Not ax Is Nothing Then
                        ax.HasTitle = True
                        ax.AxisTitle.text = parts(3)
                    End If
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If

        Case "CHART_TYPE"
            If UBound(parts) >= 2 Then
                Dim chIdx4 As Long
                chIdx4 = GetChartIndex(ws, parts(1))
                If chIdx4 > 0 And chIdx4 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx4).Chart.chartType = GetChartType(parts(2))
                End If
                ExecuteSingleCommand = True
            End If

        Case "CHART_MOVE"
            If UBound(parts) >= 2 Then
                Dim chIdx5 As Long
                chIdx5 = GetChartIndex(ws, parts(1))
                If chIdx5 > 0 And chIdx5 <= ws.ChartObjects.Count Then
                    ' Проверяем, адрес ячейки или координаты
                    On Error Resume Next
                    Dim moveCell As Range
                    Set moveCell = ws.Range(parts(2))
                    If Not moveCell Is Nothing Then
                        ws.ChartObjects(chIdx5).Left = moveCell.Left
                        ws.ChartObjects(chIdx5).Top = moveCell.Top
                    ElseIf UBound(parts) >= 3 Then
                        ws.ChartObjects(chIdx5).Left = CLng(parts(2))
                        ws.ChartObjects(chIdx5).Top = CLng(parts(3))
                    End If
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If

        Case "CHART_RESIZE"
            If UBound(parts) >= 3 Then
                Dim chIdx6 As Long
                chIdx6 = GetChartIndex(ws, parts(1))
                If chIdx6 > 0 And chIdx6 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx6).Width = CLng(parts(2))
                    ws.ChartObjects(chIdx6).Height = CLng(parts(3))
                End If
                ExecuteSingleCommand = True
            End If

        Case "CHART_DELETE"
            If UBound(parts) >= 1 Then
                Dim chIdx7 As Long
                chIdx7 = GetChartIndex(ws, parts(1))
                If chIdx7 > 0 And chIdx7 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx7).Delete
                End If
                ExecuteSingleCommand = True
            End If

        Case "CHART_DELETE_ALL"
            Dim co As ChartObject
            For Each co In ws.ChartObjects
                co.Delete
            Next co
            ExecuteSingleCommand = True

        Case "MOVE_CHART"
            ' Для совместимости - пропускаем
            ExecuteSingleCommand = True

        ' ========== СВОДНЫЕ ТАБЛИЦЫ ==========

        Case "CREATE_PIVOT"
            ' CREATE_PIVOT|источник|назначение|имя
            ' Пример: CREATE_PIVOT|A1:D10|F1|МояСводная
            ' Или: CREATE_PIVOT|Лист1!A1:D10|Лист2!A1|МояСводная
            If UBound(parts) >= 3 Then
                Dim pivotCache As pivotCache
                Dim pivotTable As pivotTable
                Dim srcRng As Range
                Dim destCell As Range
                Dim destSheet As Worksheet
                Dim pivotName As String

                ' Используем Application.Range для поддержки ссылок с именами листов
                On Error Resume Next
                Set srcRng = Application.Range(parts(1))
                If srcRng Is Nothing Then
                    ' Попробуем без имени листа
                    Set srcRng = ws.Range(parts(1))
                End If

                ' Для назначения - проверяем, нужно ли создать новый лист
                Set destCell = Application.Range(parts(2))
                If destCell Is Nothing Then
                    ' Если лист не существует, создаём его
                    Dim destParts() As String
                    If InStr(parts(2), "!") > 0 Then
                        destParts = Split(parts(2), "!")
                        Dim sheetName As String
                        sheetName = Replace(Replace(destParts(0), "'", ""), "!", "")
                        ' Проверяем существование листа
                        Dim sheetExists As Boolean
                        sheetExists = False
                        Dim wsCheck As Worksheet
                        For Each wsCheck In ActiveWorkbook.Worksheets
                            If wsCheck.Name = sheetName Then
                                sheetExists = True
                                Set destSheet = wsCheck
                                Exit For
                            End If
                        Next wsCheck
                        If Not sheetExists Then
                            Set destSheet = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
                            destSheet.Name = sheetName
                        End If
                        Set destCell = destSheet.Range(destParts(1))
                    Else
                        Set destCell = ws.Range(parts(2))
                    End If
                End If
                On Error GoTo ErrorHandler

                If Not srcRng Is Nothing And Not destCell Is Nothing Then
                    pivotName = Trim(parts(3))

                    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
                        SourceType:=xlDatabase, _
                        SourceData:=srcRng)

                    Set pivotTable = pivotCache.CreatePivotTable( _
                        TableDestination:=destCell, _
                        tableName:=pivotName)

                    ExecuteSingleCommand = True
                End If
            End If

        Case "PIVOT_ADD_ROW"
            If UBound(parts) >= 2 Then
                Dim pt As pivotTable
                Set pt = FindPivotTable(parts(1))
                If Not pt Is Nothing Then
                    On Error Resume Next
                    pt.PivotFields(parts(2)).Orientation = xlRowField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If

        Case "PIVOT_ADD_COLUMN"
            If UBound(parts) >= 2 Then
                Dim pt2 As pivotTable
                Set pt2 = FindPivotTable(parts(1))
                If Not pt2 Is Nothing Then
                    On Error Resume Next
                    pt2.PivotFields(parts(2)).Orientation = xlColumnField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If

        Case "PIVOT_ADD_VALUE"
            If UBound(parts) >= 3 Then
                Dim pt3 As pivotTable
                Dim pfFunc As Long
                Set pt3 = FindPivotTable(parts(1))
                If Not pt3 Is Nothing Then
                    Select Case UCase(Trim(parts(3)))
                        Case "SUM": pfFunc = xlSum
                        Case "COUNT": pfFunc = xlCount
                        Case "AVERAGE": pfFunc = xlAverage
                        Case "MAX": pfFunc = xlMax
                        Case "MIN": pfFunc = xlMin
                        Case Else: pfFunc = xlSum
                    End Select
                    On Error Resume Next
                    pt3.AddDataField pt3.PivotFields(parts(2)), , pfFunc
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If

        Case "PIVOT_ADD_FILTER"
            If UBound(parts) >= 2 Then
                Dim pt4 As pivotTable
                Set pt4 = FindPivotTable(parts(1))
                If Not pt4 Is Nothing Then
                    On Error Resume Next
                    pt4.PivotFields(parts(2)).Orientation = xlPageField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If

        Case "PIVOT_REFRESH"
            If UBound(parts) >= 1 Then
                Dim pt5 As pivotTable
                Set pt5 = FindPivotTable(parts(1))
                If Not pt5 Is Nothing Then
                    pt5.RefreshTable
                End If
                ExecuteSingleCommand = True
            End If

        Case "PIVOT_REFRESH_ALL"
            ActiveWorkbook.RefreshAll
            ExecuteSingleCommand = True

        ' ========== ЛИСТЫ ==========

        Case "ADD_SHEET"
            If UBound(parts) >= 1 Then
                Dim newSheet As Worksheet
                Set newSheet = ActiveWorkbook.Worksheets.Add
                newSheet.Name = parts(1)
                ExecuteSingleCommand = True
            End If

        Case "ADD_SHEET_AFTER"
            If UBound(parts) >= 2 Then
                Dim newSheet2 As Worksheet
                Set newSheet2 = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(parts(2)))
                newSheet2.Name = parts(1)
                ExecuteSingleCommand = True
            End If

        Case "DELETE_SHEET"
            If UBound(parts) >= 1 Then
                Application.DisplayAlerts = False
                ActiveWorkbook.Worksheets(parts(1)).Delete
                Application.DisplayAlerts = True
                ExecuteSingleCommand = True
            End If

        Case "RENAME_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Name = parts(2)
                ExecuteSingleCommand = True
            End If

        Case "COPY_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Copy after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
                ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count).Name = parts(2)
                ExecuteSingleCommand = True
            End If

        Case "MOVE_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Move before:=ActiveWorkbook.Worksheets(CLng(parts(2)))
                ExecuteSingleCommand = True
            End If

        Case "HIDE_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Visible = xlSheetHidden
                ExecuteSingleCommand = True
            End If

        Case "SHOW_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Visible = xlSheetVisible
                ExecuteSingleCommand = True
            End If

        Case "ACTIVATE_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Activate
                ExecuteSingleCommand = True
            End If

        Case "TAB_COLOR"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Tab.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "PROTECT_SHEET"
            If UBound(parts) >= 1 Then
                Dim pwd As String
                pwd = ""
                If UBound(parts) >= 2 Then pwd = parts(2)
                ActiveWorkbook.Worksheets(parts(1)).Protect Password:=pwd
                ExecuteSingleCommand = True
            End If

        Case "UNPROTECT_SHEET"
            If UBound(parts) >= 1 Then
                Dim pwd2 As String
                pwd2 = ""
                If UBound(parts) >= 2 Then pwd2 = parts(2)
                ActiveWorkbook.Worksheets(parts(1)).Unprotect Password:=pwd2
                ExecuteSingleCommand = True
            End If

        ' ========== ИМЕНОВАННЫЕ ДИАПАЗОНЫ ==========

        Case "CREATE_NAME"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Names.Add Name:=parts(1), RefersTo:=ws.Range(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "DELETE_NAME"
            If UBound(parts) >= 1 Then
                On Error Resume Next
                ActiveWorkbook.Names(parts(1)).Delete
                On Error GoTo ErrorHandler
                ExecuteSingleCommand = True
            End If

        ' ========== УСЛОВНОЕ ФОРМАТИРОВАНИЕ ==========

        Case "COND_HIGHLIGHT"
            ' COND_HIGHLIGHT|диапазон|формула|цвет (4 части)
            ' COND_HIGHLIGHT|диапазон|оператор|значение|цвет (5 частей)
            If UBound(parts) >= 3 Then
                Dim hlRng As Range
                Dim hlFormula As String
                Dim hlFirstCell As String
                Dim hlColor As String
                Dim hlOp As String
                Dim hlFC As Object

                Debug.Print "COND_HIGHLIGHT: Step 1 - parsing range: " & parts(1)
                Set hlRng = ws.Range(parts(1))
                Debug.Print "COND_HIGHLIGHT: Step 2 - range set OK: " & hlRng.address
                hlFirstCell = hlRng.Cells(1, 1).address(False, False)
                Debug.Print "COND_HIGHLIGHT: Step 3 - firstCell: " & hlFirstCell

                If UBound(parts) = 3 Then
                    ' 4 части: диапазон|формула|цвет
                    Debug.Print "COND_HIGHLIGHT: Step 4a - 4 parts mode"
                    hlFormula = Trim(parts(2))
                    hlColor = Trim(parts(3))
                Else
                    ' 5 частей: диапазон|оператор|значение|цвет
                    Debug.Print "COND_HIGHLIGHT: Step 4b - 5 parts mode"
                    hlOp = Trim(parts(2))
                    hlColor = Trim(parts(4))
                    Debug.Print "COND_HIGHLIGHT: Step 5 - op=" & hlOp & " color=" & hlColor

                    Select Case hlOp
                        Case ">", "<", ">=", "<=", "<>"
                            hlFormula = hlFirstCell & hlOp & Trim(parts(3))
                        Case "="
                            If IsNumeric(Trim(parts(3))) Then
                                hlFormula = hlFirstCell & "=" & Trim(parts(3))
                            Else
                                hlFormula = Trim(parts(3))
                            End If
                        Case Else
                            hlFormula = Trim(parts(3))
                    End Select
                End If

                Debug.Print "COND_HIGHLIGHT: Step 6 - formula before =: " & hlFormula

                ' Добавляем = в начало если нет
                If Left(hlFormula, 1) <> "=" Then hlFormula = "=" & hlFormula

                Debug.Print "COND_HIGHLIGHT: Step 7 - formula after =: " & hlFormula

                ' Локализуем формулу (MOD -> ОСТАТ и т.д., запятая -> точка с запятой)
                hlFormula = LocalizeFormula(hlFormula)
                Debug.Print "COND_HIGHLIGHT: Step 8 - formula localized: " & hlFormula
                Debug.Print "COND_HIGHLIGHT: Step 9 - adding FormatCondition..."

                ' ВАЖНО: Выбираем первую ячейку диапазона, чтобы Excel правильно
                ' интерпретировал относительные ссылки в формуле
                hlRng.Cells(1, 1).Select

                ' Добавляем условное форматирование
                Set hlFC = hlRng.FormatConditions.Add(Type:=xlExpression, Formula1:=hlFormula)

                Debug.Print "COND_HIGHLIGHT: Step 10 - hlColor=[" & hlColor & "]"

                Dim hlColorValue As Long
                On Error Resume Next
                hlColorValue = ParseColor(hlColor)
                If Err.Number <> 0 Then
                    Debug.Print "COND_HIGHLIGHT: ParseColor ERROR: " & Err.Number & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

                Debug.Print "COND_HIGHLIGHT: Step 10a - colorValue=" & hlColorValue
                Debug.Print "COND_HIGHLIGHT: Step 10b - hlFC type=" & TypeName(hlFC)

                hlFC.Interior.Color = hlColorValue

                Debug.Print "COND_HIGHLIGHT: Step 11 - DONE"
                ExecuteSingleCommand = True
            End If

        Case "COND_TOP"
            If UBound(parts) >= 3 Then
                Dim cfRng2 As Range
                Set cfRng2 = ws.Range(parts(1))
                cfRng2.FormatConditions.AddTop10
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).TopBottom = xlTop10Top
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).Rank = CLng(parts(2))
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If

        Case "COND_BOTTOM"
            If UBound(parts) >= 3 Then
                Dim cfRng3 As Range
                Set cfRng3 = ws.Range(parts(1))
                cfRng3.FormatConditions.AddTop10
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).TopBottom = xlTop10Bottom
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).Rank = CLng(parts(2))
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If

        Case "COND_DUPLICATE"
            If UBound(parts) >= 2 Then
                Dim cfRng4 As Range
                Set cfRng4 = ws.Range(parts(1))
                cfRng4.FormatConditions.AddUniqueValues
                cfRng4.FormatConditions(cfRng4.FormatConditions.Count).DupeUnique = xlDuplicate
                cfRng4.FormatConditions(cfRng4.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "COND_UNIQUE"
            If UBound(parts) >= 2 Then
                Dim cfRng5 As Range
                Set cfRng5 = ws.Range(parts(1))
                cfRng5.FormatConditions.AddUniqueValues
                cfRng5.FormatConditions(cfRng5.FormatConditions.Count).DupeUnique = xlUnique
                cfRng5.FormatConditions(cfRng5.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "COND_TEXT"
            If UBound(parts) >= 3 Then
                Dim cfRng6 As Range
                Set cfRng6 = ws.Range(parts(1))
                cfRng6.FormatConditions.Add Type:=xlTextString, String:=parts(2), TextOperator:=xlContains
                cfRng6.FormatConditions(cfRng6.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If

        Case "COND_BLANK"
            If UBound(parts) >= 2 Then
                Dim cfRng7 As Range
                Set cfRng7 = ws.Range(parts(1))
                cfRng7.FormatConditions.Add Type:=xlBlanksCondition
                cfRng7.FormatConditions(cfRng7.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "COND_NOT_BLANK"
            If UBound(parts) >= 2 Then
                Dim cfRng8 As Range
                Set cfRng8 = ws.Range(parts(1))
                cfRng8.FormatConditions.Add Type:=xlNoBlanksCondition
                cfRng8.FormatConditions(cfRng8.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "DATA_BARS"
            If UBound(parts) >= 2 Then
                Dim cfRng9 As Range
                Set cfRng9 = ws.Range(parts(1))
                cfRng9.FormatConditions.AddDatabar
                cfRng9.FormatConditions(cfRng9.FormatConditions.Count).BarColor.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If

        Case "COLOR_SCALE"
            If UBound(parts) >= 3 Then
                Dim cfRng10 As Range
                Set cfRng10 = ws.Range(parts(1))
                cfRng10.FormatConditions.AddColorScale ColorScaleType:=2
                cfRng10.FormatConditions(cfRng10.FormatConditions.Count).ColorScaleCriteria(1).FormatColor.Color = ParseColor(parts(2))
                cfRng10.FormatConditions(cfRng10.FormatConditions.Count).ColorScaleCriteria(2).FormatColor.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If

        Case "COLOR_SCALE3"
            If UBound(parts) >= 4 Then
                Dim cfRng11 As Range
                Set cfRng11 = ws.Range(parts(1))
                cfRng11.FormatConditions.AddColorScale ColorScaleType:=3
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(1).FormatColor.Color = ParseColor(parts(2))
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(2).FormatColor.Color = ParseColor(parts(3))
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(3).FormatColor.Color = ParseColor(parts(4))
                ExecuteSingleCommand = True
            End If

        Case "ICON_SET"
            If UBound(parts) >= 2 Then
                Dim cfRng12 As Range
                Dim iconSetType As Long
                Set cfRng12 = ws.Range(parts(1))

                Select Case UCase(Trim(parts(2)))
                    Case "ARROWS": iconSetType = 1 ' xl3Arrows
                    Case "FLAGS": iconSetType = 7 ' xl3Flags
                    Case "STARS": iconSetType = 13 ' xl3Stars
                    Case "BARS": iconSetType = 14 ' xl4RedToBlack
                    Case Else: iconSetType = 1
                End Select

                cfRng12.FormatConditions.AddIconSetCondition
                cfRng12.FormatConditions(cfRng12.FormatConditions.Count).IconSet = ActiveWorkbook.IconSets(iconSetType)
                ExecuteSingleCommand = True
            End If

        Case "CLEAR_COND_FORMAT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).FormatConditions.Delete
                ExecuteSingleCommand = True
            End If

        ' ========== ПРОВЕРКА ДАННЫХ ==========

        Case "VALIDATION_LIST"
            If UBound(parts) >= 2 Then
                Dim valRng As Range
                Set valRng = ws.Range(parts(1))
                valRng.Validation.Delete
                valRng.Validation.Add Type:=xlValidateList, Formula1:=Replace(parts(2), ";", ",")
                ExecuteSingleCommand = True
            End If

        Case "VALIDATION_NUMBER"
            If UBound(parts) >= 3 Then
                Dim valRng2 As Range
                Set valRng2 = ws.Range(parts(1))
                valRng2.Validation.Delete
                valRng2.Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If

        Case "VALIDATION_DATE"
            If UBound(parts) >= 3 Then
                Dim valRng3 As Range
                Set valRng3 = ws.Range(parts(1))
                valRng3.Validation.Delete
                valRng3.Validation.Add Type:=xlValidateDate, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If

        Case "VALIDATION_TEXT_LENGTH"
            If UBound(parts) >= 3 Then
                Dim valRng4 As Range
                Set valRng4 = ws.Range(parts(1))
                valRng4.Validation.Delete
                valRng4.Validation.Add Type:=xlValidateTextLength, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If

        Case "VALIDATION_CUSTOM"
            If UBound(parts) >= 2 Then
                Dim valRng5 As Range
                Dim valFormula As String
                Set valRng5 = ws.Range(parts(1))
                valFormula = LocalizeFormula(parts(2))
                valRng5.Validation.Delete
                valRng5.Validation.Add Type:=xlValidateCustom, Formula1:=valFormula
                ExecuteSingleCommand = True
            End If

        Case "CLEAR_VALIDATION"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Validation.Delete
                ExecuteSingleCommand = True
            End If

        ' ========== КОММЕНТАРИИ ==========

        Case "ADD_COMMENT"
            If UBound(parts) >= 2 Then
                Dim cmtCell As Range
                Set cmtCell = ws.Range(parts(1))
                If Not cmtCell.Comment Is Nothing Then cmtCell.Comment.Delete
                cmtCell.AddComment parts(2)
                ExecuteSingleCommand = True
            End If

        Case "EDIT_COMMENT"
            If UBound(parts) >= 2 Then
                Dim cmtCell2 As Range
                Set cmtCell2 = ws.Range(parts(1))
                If Not cmtCell2.Comment Is Nothing Then
                    cmtCell2.Comment.text text:=parts(2)
                End If
                ExecuteSingleCommand = True
            End If

        Case "DELETE_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell3 As Range
                Set cmtCell3 = ws.Range(parts(1))
                If Not cmtCell3.Comment Is Nothing Then cmtCell3.Comment.Delete
                ExecuteSingleCommand = True
            End If

        Case "SHOW_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell4 As Range
                Set cmtCell4 = ws.Range(parts(1))
                If Not cmtCell4.Comment Is Nothing Then cmtCell4.Comment.Visible = True
                ExecuteSingleCommand = True
            End If

        Case "HIDE_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell5 As Range
                Set cmtCell5 = ws.Range(parts(1))
                If Not cmtCell5.Comment Is Nothing Then cmtCell5.Comment.Visible = False
                ExecuteSingleCommand = True
            End If

        Case "SHOW_ALL_COMMENTS"
            Dim cmt As Comment
            For Each cmt In ws.Comments
                cmt.Visible = True
            Next cmt
            ExecuteSingleCommand = True

        Case "HIDE_ALL_COMMENTS"
            Dim cmt2 As Comment
            For Each cmt2 In ws.Comments
                cmt2.Visible = False
            Next cmt2
            ExecuteSingleCommand = True

        ' ========== ГИПЕРССЫЛКИ ==========

        Case "ADD_HYPERLINK"
            If UBound(parts) >= 3 Then
                ws.Hyperlinks.Add Anchor:=ws.Range(parts(1)), address:=parts(2), TextToDisplay:=parts(3)
                ExecuteSingleCommand = True
            End If

        Case "ADD_HYPERLINK_CELL"
            If UBound(parts) >= 3 Then
                ws.Hyperlinks.Add Anchor:=ws.Range(parts(1)), address:="", SubAddress:=parts(2), TextToDisplay:=parts(3)
                ExecuteSingleCommand = True
            End If

        Case "REMOVE_HYPERLINK"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Hyperlinks.Delete
                ExecuteSingleCommand = True
            End If

        ' ========== ЗАЩИТА ==========

        Case "LOCK_CELLS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Locked = True
                ExecuteSingleCommand = True
            End If

        Case "UNLOCK_CELLS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Locked = False
                ExecuteSingleCommand = True
            End If

        ' ========== ОБЛАСТЬ ПРОСМОТРА ==========

        Case "FREEZE_PANES"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Select
                ActiveWindow.FreezePanes = True
                ExecuteSingleCommand = True
            End If

        Case "FREEZE_TOP_ROW"
            ws.Range("A2").Select
            ActiveWindow.FreezePanes = True
            ExecuteSingleCommand = True

        Case "FREEZE_FIRST_COLUMN"
            ws.Range("B1").Select
            ActiveWindow.FreezePanes = True
            ExecuteSingleCommand = True

        Case "UNFREEZE_PANES"
            ActiveWindow.FreezePanes = False
            ExecuteSingleCommand = True

        Case "ZOOM"
            If UBound(parts) >= 1 Then
                ActiveWindow.Zoom = CLng(parts(1))
                ExecuteSingleCommand = True
            End If

        Case "GOTO"
            If UBound(parts) >= 1 Then
                Application.Goto Reference:=ws.Range(parts(1)), Scroll:=True
                ExecuteSingleCommand = True
            End If

        Case "SELECT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Select
                ExecuteSingleCommand = True
            End If

        ' ========== ПЕЧАТЬ ==========

        Case "SET_PRINT_AREA"
            If UBound(parts) >= 1 Then
                ws.PageSetup.PrintArea = parts(1)
                ExecuteSingleCommand = True
            End If

        Case "CLEAR_PRINT_AREA"
            ws.PageSetup.PrintArea = ""
            ExecuteSingleCommand = True

        Case "PAGE_ORIENTATION"
            If UBound(parts) >= 1 Then
                If UCase(Trim(parts(1))) = "LANDSCAPE" Then
                    ws.PageSetup.Orientation = xlLandscape
                Else
                    ws.PageSetup.Orientation = xlPortrait
                End If
                ExecuteSingleCommand = True
            End If

        Case "PAGE_MARGINS"
            If UBound(parts) >= 4 Then
                With ws.PageSetup
                    .LeftMargin = Application.CentimetersToPoints(CDbl(parts(1)))
                    .RightMargin = Application.CentimetersToPoints(CDbl(parts(2)))
                    .TopMargin = Application.CentimetersToPoints(CDbl(parts(3)))
                    .BottomMargin = Application.CentimetersToPoints(CDbl(parts(4)))
                End With
                ExecuteSingleCommand = True
            End If

        Case "PRINT_TITLES_ROWS"
            If UBound(parts) >= 2 Then
                ws.PageSetup.PrintTitleRows = "$" & parts(1) & ":$" & parts(2)
                ExecuteSingleCommand = True
            End If

        Case "PRINT_TITLES_COLS"
            If UBound(parts) >= 2 Then
                ws.PageSetup.PrintTitleColumns = "$" & parts(1) & ":$" & parts(2)
                ExecuteSingleCommand = True
            End If

        Case "PRINT_GRIDLINES"
            If UBound(parts) >= 1 Then
                ws.PageSetup.PrintGridlines = (UCase(Trim(parts(1))) = "TRUE")
                ExecuteSingleCommand = True
            End If

        Case "FIT_TO_PAGE"
            If UBound(parts) >= 2 Then
                With ws.PageSetup
                    .Zoom = False
                    .FitToPagesWide = CLng(parts(1))
                    .FitToPagesTall = CLng(parts(2))
                End With
                ExecuteSingleCommand = True
            End If

        ' ========== ИЗОБРАЖЕНИЯ ==========

        Case "INSERT_PICTURE"
            If UBound(parts) >= 5 Then
                Dim pic As Object
                Set pic = ws.Shapes.AddPicture(parts(1), msoFalse, msoTrue, _
                    CLng(parts(2)), CLng(parts(3)), CLng(parts(4)), CLng(parts(5)))
                ExecuteSingleCommand = True
            End If

        Case "DELETE_PICTURES"
            Dim shp As Shape
            For Each shp In ws.Shapes
                If shp.Type = msoPicture Then shp.Delete
            Next shp
            ExecuteSingleCommand = True

        ' ========== ФОРМЫ ==========

        Case "ADD_BUTTON"
            If UBound(parts) >= 5 Then
                Dim btn As Object
                Set btn = ws.Buttons.Add(CLng(parts(1)), CLng(parts(2)), CLng(parts(3)), CLng(parts(4)))
                btn.Caption = parts(5)
                ExecuteSingleCommand = True
            End If

        Case "ADD_CHECKBOX"
            If UBound(parts) >= 2 Then
                Dim chk As Object
                Dim chkCell As Range
                Set chkCell = ws.Range(parts(1))
                Set chk = ws.CheckBoxes.Add(chkCell.Left, chkCell.Top, 100, 15)
                chk.Caption = parts(2)
                ExecuteSingleCommand = True
            End If

        Case "ADD_DROPDOWN"
            If UBound(parts) >= 2 Then
                Dim dd As Object
                Dim ddCell As Range
                Set ddCell = ws.Range(parts(1))
                Set dd = ws.DropDowns.Add(ddCell.Left, ddCell.Top, 100, 15)
                dd.List = Split(parts(2), ";")
                ExecuteSingleCommand = True
            End If

        Case "DELETE_SHAPES"
            Dim shp2 As Shape
            For Each shp2 In ws.Shapes
                shp2.Delete
            Next shp2
            ExecuteSingleCommand = True

        ' ========== СПЕЦИАЛЬНЫЕ ==========

        Case "CALCULATE"
            Application.Calculate
            ExecuteSingleCommand = True

        Case "CALCULATE_SHEET"
            ws.Calculate
            ExecuteSingleCommand = True

        Case "TEXT_TO_COLUMNS"
            If UBound(parts) >= 2 Then
                Dim ttcRng As Range
                Dim delim As String
                Set ttcRng = ws.Range(parts(1))
                delim = parts(2)

                Dim delimTab As Boolean, delimSemi As Boolean, delimComma As Boolean, delimSpace As Boolean, delimOther As Boolean
                Dim otherChar As String

                Select Case UCase(delim)
                    Case "TAB": delimTab = True
                    Case "SEMICOLON", ";": delimSemi = True
                    Case "COMMA", ",": delimComma = True
                    Case "SPACE", " ": delimSpace = True
                    Case Else
                        delimOther = True
                        otherChar = delim
                End Select

                ttcRng.TextToColumns Destination:=ttcRng, DataType:=xlDelimited, _
                    Tab:=delimTab, Semicolon:=delimSemi, Comma:=delimComma, _
                    Space:=delimSpace, Other:=delimOther, otherChar:=otherChar

                ExecuteSingleCommand = True
            End If

        Case "REMOVE_SPACES"
            If UBound(parts) >= 1 Then
                Dim spRng As Range, spCell As Range
                Set spRng = ws.Range(parts(1))
                For Each spCell In spRng
                    If Not IsEmpty(spCell.value) Then
                        spCell.value = Application.WorksheetFunction.Trim(spCell.value)
                    End If
                Next spCell
                ExecuteSingleCommand = True
            End If

        Case "UPPER_CASE"
            If UBound(parts) >= 1 Then
                Dim ucRng As Range, ucCell As Range
                Set ucRng = ws.Range(parts(1))
                For Each ucCell In ucRng
                    If Not IsEmpty(ucCell.value) Then
                        ucCell.value = UCase(ucCell.value)
                    End If
                Next ucCell
                ExecuteSingleCommand = True
            End If

        Case "LOWER_CASE"
            If UBound(parts) >= 1 Then
                Dim lcRng As Range, lcCell As Range
                Set lcRng = ws.Range(parts(1))
                For Each lcCell In lcRng
                    If Not IsEmpty(lcCell.value) Then
                        lcCell.value = LCase(lcCell.value)
                    End If
                Next lcCell
                ExecuteSingleCommand = True
            End If

        Case "PROPER_CASE"
            If UBound(parts) >= 1 Then
                Dim pcRng As Range, pcCell As Range
                Set pcRng = ws.Range(parts(1))
                For Each pcCell In pcRng
                    If Not IsEmpty(pcCell.value) Then
                        pcCell.value = Application.WorksheetFunction.Proper(pcCell.value)
                    End If
                Next pcCell
                ExecuteSingleCommand = True
            End If

        Case "FLASH_FILL"
            If UBound(parts) >= 1 Then
                On Error Resume Next
                ws.Range(parts(1)).FlashFill
                On Error GoTo ErrorHandler
                ExecuteSingleCommand = True
            End If

        Case "SUBTOTAL"
            If UBound(parts) >= 3 Then
                Dim stRng As Range
                Dim stFunc As Long
                Dim stCol As Long
                Set stRng = ws.Range(parts(1))

                Select Case UCase(Trim(parts(2)))
                    Case "SUM": stFunc = xlSum
                    Case "COUNT": stFunc = xlCount
                    Case "AVERAGE": stFunc = xlAverage
                    Case "MAX": stFunc = xlMax
                    Case "MIN": stFunc = xlMin
                    Case Else: stFunc = xlSum
                End Select

                stCol = GetColumnNumber(parts(3), stRng, ws)
                stRng.Subtotal GroupBy:=1, Function:=stFunc, TotalList:=Array(stCol)
                ExecuteSingleCommand = True
            End If

        Case "REMOVE_SUBTOTALS"
            ws.UsedRange.RemoveSubtotal
            ExecuteSingleCommand = True

        Case Else
            ExecuteSingleCommand = False
    End Select

    Exit Function

ErrorHandler:
    Debug.Print "ExecuteSingleCommand Error: " & Err.Number & " - " & Err.Description & " | CMD: " & cmd
    ExecuteSingleCommand = False
End Function
