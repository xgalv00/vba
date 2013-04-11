Attribute VB_Name = "MasterBPC_tmp"
''*************************
'' Набор макросов для ВРС
''*************************
'' Автор: Александр Мастер
'' Автор: Алексей Воронин
'' Автор: Владимир Аниканов
''*************************
'' Внесены небольшие изменения 14.08.2012
'' Дата:  11.07.2012
''*************************
'
'Public sh           As Worksheet
'Public ActCell      As Range
'Public EVDRE()      As Variant
'Public Event_Name   As String
'Dim InProcess       As Boolean
'Dim IsRegistered    As Boolean
'Dim myConnection    As Object
'Dim sap, fm, tabl
'Const UsingLocks    As Boolean = False
'
'    'Application.Run "MNU_eTOOLS_REFRESH" 'Запрос перед обновлением
'    'Application.Run "MNU_eSUBMIT_REFRESH" 'Отправка данных с запросом и обновление
'
'    'Application.Run "MNU_eSUBMIT_REFSCHEDULE_SHEET_REFRESH" 'Отправка данных без запроса и обновление
'    'Application.Run "MNU_eSUBMIT_REFSCHEDULE_SHEET_NOACTION" 'Отправка данных без запросов и без обновлений
'    'Application.Run "MNU_eSUBMIT_REFSCHEDULE_SHEET_CLEARANDREFRESH" 'Отправка данных без запроса и обновление
'
'    'Application.Run "MNU_eSUBMIT_REFSCHEDULE_BOOK_REFRESH"
'    'Application.Run "MNU_eSUBMIT_REFSCHEDULE_BOOK_NOACTION"
'    'Application.Run "MNU_eSUBMIT_REFSCHEDULE_BOOK_CLEARANDREFRESH"
'    'Application.Run "MNU_eSUBMIT_REFSCHEDULE_BOOK_NOACTION_SHOWRESULT"
'
''************************************************************************************
''************************************************************************************
''************************************************************************************
'
'Dim cvNotMatch As Boolean
'
'Public Sub onWorkBookActivate()
'    'This sub will be called on workbook activate
'    Dim tmpStr As String
'
'    'check for connection
'    On Error Resume Next
'    tmpStr = Evaluate("EVAST()")
'    If Err.Number <> 0 Then
'        Exit Sub
'    End If
'    On Error GoTo 0
'    'additional check
'    If tmpStr = "DTEK" Then
'        Call generateCaution
'    End If
'
'End Sub
'
''checks if cvw is changed
'Private Sub generateCaution()
'
'    Dim tmpCell As Range, tmpRow As Integer, CVWCells As Collection
'    Dim wBook As Workbook, wSheet As Worksheet, wRange As Range
'    Dim cellFormula As String
'    Dim formResult As String
'
'
'
'    cvNotMatch = False
'    Set wBook = ActiveWorkbook
'    Set wSheet = ActiveSheet
'    Set wRange = wSheet.UsedRange
'    'unprotect sheet before do smth. Sub doesn't work without this.
'    'wSheet.Unprotect Pass(wSheet)
'    Set CVWCells = findCVWCells(wRange)
'    'exit if cells not found
'    If CVWCells Is Nothing Then
'        Exit Sub
'    End If
'    'Compares each cell value with actual value of cv
'    For Each tmpCell In CVWCells
'        cellFormula = tmpCell.Formula
'        cellFormula = Right(tmpCell.Formula, Len(tmpCell.Formula) - 1)
'
'        'On Error Resume Next
'        formResult = Evaluate(cellFormula)
'
'        If tmpCell.Value <> formResult Then
'            'generates caution if cv showed in workbook mismatch with actual cv
'            cvNotMatch = True
'            Call showCaution
'            Exit Sub
'        End If
'    Next
'    'protects workbook back
'    'wSheet.Protect Pass(wSheet)
'
'End Sub
'
''finds all cells that contain evcvw function
'Private Function findCVWCells(wRange As Range) As Collection
'
'    Dim tmpCell As Range, tmpRow As Integer, tmpColl As New Collection
'    Dim firstFoundCell As Range
'
'    'look for cell with EvCvw formula in it
'    Set firstFoundCell = wRange.Find(What:="*evcvw*", LookIn:=xlFormulas, LookAt:=xlWhole)
'    'if any cell hasn't found exits
'    If firstFoundCell Is Nothing Then
'        Set findCVWCells = Nothing
'        Exit Function
'    End If
'    tmpColl.Add firstFoundCell
'    'searches for next occurence of evcvw
'    Set tmpCell = wRange.FindNext(firstFoundCell)
'
'    'loops until search wraps to first found cell
'    Do While firstFoundCell.Address <> tmpCell.Address
'        tmpColl.Add tmpCell
'
'        Set tmpCell = wRange.FindNext(tmpCell)
'    Loop
'
'    Set findCVWCells = tmpColl
'End Function
'
'Public Sub showCaution()
'    MsgBox "Текущий ракурс отличается от фильтров в рабочей книге. Смотри инструкцию", vbCritical, "Ошибка"
'End Sub
''************************************************************************************
''************************************************************************************
''************************************************************************************
'' Развернуть все
'Private Sub Expand()
'    Application.Run "MNU_eTOOLS_EXPAND"
'End Sub
'
'' Обновить книгу
'Private Sub refresh()
'    Application.Run "MNU_eTOOLS_REFRESH" 'Запрос перед обновлением
'End Sub
'
''Отправить данные без подтверждения и показать результат
'Private Sub Submit_NoAction_ShowResult()
'    Application.Run "MNU_ESUBMIT_REFSCHEDULE_BOOK_NOACTION_SHOWRESULT"
'End Sub
'
'Function BEFORE_EXPAND(argument As String)
'    Event_Name = "BEFORE_EXPAND"
'    ' Установить основные параметры для оптимизации выполнения макросов
'    SetOptimizeMode sh, ActCell
'
'    ' Есть ли колонка для обработки
'    Set f = sh.Rows(1).Find(Event_Name, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Снять объединение ячеек
'        UnMergeCells sh, f.Column
'    End If
'
'    ' Снять основные параметры для оптимизации выполнения макросов
'    SetNormalMode sh, ActCell
'    BEFORE_EXPAND = True
'End Function
'
Function AFTER_EXPAND(argument As String)
'    Event_Name = "AFTER_EXPAND"
'    ' Установить основные параметры для оптимизации выполнения макросов
'    SetOptimizeMode sh, ActCell
'
'    ' Есть ли колонка для обработки
'    Set f = sh.Rows(1).Find(Event_Name, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Обновить книгу
'        Set f2 = sh.Columns(f.Column).Find("Refresh", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'        If Not f2 Is Nothing Then
'            Application.Run "MNU_eTOOLS_REFRESH"
'            ' Установить основные параметры для оптимизации выполнения макросов
'            SetOptimizeMode sh, ActCell
'        End If
'        ' Скопировать содержимое
'        CopyPaste sh, f.Column
'        ' Применить формулы
'        ApplyFormulas sh, f.Column
'        ' Скрыть столбцы
'        HideColumns sh, f.Column
'        ' Объединить ячейки
'        MergeCells sh, f.Column
'    End If
'    ' Удалить строки
'    DeleteRows sh
'    ' Сортировать строки
'    Sort sh
'    ' Скрыть строки
'    HideRows sh
'    ' Отрегистрировать пользователя от всех форм
'    Locks_OffBook
'
'    ' Снять основные параметры для оптимизации выполнения макросов
'    SetNormalMode sh, ActCell

    'call for workstatus report
    Call prepareWorkspace

    AFTER_EXPAND = True
End Function
'
'Function BEFORE_REFRESH(argument As String)
'    Event_Name = "BEFORE_REFRESH"
'    ' Установить основные параметры для оптимизации выполнения макросов
'    'SetOptimizeMode sh, ActCell
'
'    ' Снять основные параметры для оптимизации выполнения макросов
'    'SetNormalMode sh, ActCell
'    BEFORE_REFRESH = True
'End Function
'
Function AFTER_REFRESH(argument As String)
'    Event_Name = "AFTER_REFRESH"
'    ' Установить основные параметры для оптимизации выполнения макросов
'    SetOptimizeMode sh, ActCell
'
'    ' Есть ли колонка для обработки
'    Set f = sh.Rows(1).Find(Event_Name, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Применить формулы
'        ApplyFormulas sh, f.Column
'        ' Развернуть все
'        Set f2 = sh.Columns(f.Column).Find("Expand", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'        If Not f2 Is Nothing Then
'            Application.Run "MNU_eTOOLS_EXPAND"
'            ' Установить основные параметры для оптимизации выполнения макросов
'            SetOptimizeMode sh, ActCell
'        End If
'    End If
'    ' Отрегистрировать пользователя от всех форм
'    'Locks_OffBook
'    'Call prepareWorkspace
'
'    ' Снять основные параметры для оптимизации выполнения макросов
'    SetNormalMode sh, ActCell

    'call for workstatus report
    Call prepareWorkspace

    AFTER_REFRESH = True
End Function
'
'Function BEFORE_SEND(argument As String)
'    Event_Name = "BEFORE_SEND"
'    If cvNotMatch Then
'        Call showCaution
'        Exit Function
'    End If
'    ' Установить основные параметры для оптимизации выполнения макросов
'    SetOptimizeMode sh, ActCell
'
'    ' Есть ли колонка для обработки
'    Set f = sh.Rows(1).Find(Event_Name, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Скопировать содержимое
'        CopyPaste sh, f.Column
'        ' Применить формулы
'        ApplyFormulas sh, f.Column
'        ' Обновить книгу
'        Set f2 = sh.Columns(f.Column).Find("Refresh", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'        If Not f2 Is Nothing Then
'            Application.Run "MNU_eTOOLS_REFRESH"
'            ' Установить основные параметры для оптимизации выполнения макросов
'            SetOptimizeMode sh, ActCell
'        End If
'    End If
'    ' Установить рабочие статусы
'    'SetWorkStatus
'
'    ' Есть ли ошибки на листе (проверки)
'    If Check(sh) > 0 Then
'        BEFORE_SEND = False
'    Else
'        BEFORE_SEND = True
'    End If
'
'    ' Снять основные параметры для оптимизации выполнения макросов
'    SetNormalMode sh, ActCell
'End Function
'
'Function AFTER_SEND(argument As String)
'    Event_Name = "AFTER_SEND"
'    ' Установить основные параметры для оптимизации выполнения макросов
'    SetOptimizeMode sh, ActCell
'
'    ' Есть ли колонка для обработки
'    Set f = sh.Rows(1).Find(Event_Name, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Развернуть все
'        Set f2 = sh.Columns(f.Column).Find("Expand", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'        If Not f2 Is Nothing Then
'            Application.Run "MNU_eTOOLS_EXPAND"
'            ' Установить основные параметры для оптимизации выполнения макросов
'            SetOptimizeMode sh, ActCell
'        End If
'        ' Обновить книгу
'        Set f2 = sh.Columns(f.Column).Find("Refresh", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'        If Not f2 Is Nothing Then
'            Application.Run "MNU_eTOOLS_REFRESH"
'            ' Установить основные параметры для оптимизации выполнения макросов
'            SetOptimizeMode sh, ActCell
'        End If
'    End If
'    ' Отрегистрировать пользователя от одной формы
'    Locks_OffSheet sh
'    'Locks_Off sh
'
'    ' Снять основные параметры для оптимизации выполнения макросов
'    SetNormalMode sh, ActCell
'    AFTER_SEND = True
'End Function
'
'Private Function Check(sh)
''
'' Проверить наличие ошибок на листе
''
'    Check = 0
'    ' Поиск столбца-фильтра
'    Set f = sh.Rows(1).Find("Check", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        Msg = "На листе есть ошибки. Вы уверены что хотите сохранить данные?"
'        Style = vbYesNo + vbQuestion + vbDefaultButton2
'        Title = "Сохранение данных"
'        ' Если есть ошибки, то вывести сообщение
'        If f.Offset(0, 1).Value > 0 Then
'            mb = MsgBox(Msg, Style, Title)
'        End If
'        If mb = vbNo Then Check = 1
'    End If
'End Function
'
'Public Function Pass(sh)
''
'' Прочитать пароль на листе
''
'    ' Поиск ячейки-маркера
'    Set f = sh.Cells.Find("PasswordBPC", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        Set f = sh.Cells(f.Row + 1, f.Column)
'        Pass = f.Value
'        If sh.ProtectContents = False Then
'            f.NumberFormat = ";;;"
'            f.Locked = True
'            f.FormulaHidden = True
'            Set r = Range(sh.Cells(f.Row - 1, f.Column), f)
'            r.Interior.ThemeColor = xlThemeColorAccent1
'            r.Interior.TintAndShade = 0.4
'        End If
'    End If
'End Function
'
'Private Sub SetOptimizeMode(sh, ActCell)
''
'' Установить основные параметры для оптимизации выполнения макросов
''
'    Application.ReferenceStyle = xlA1
'    Application.EnableEvents = False
'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False
'    ' Запомнить позицию курсора
'    Set sh = ActiveWorkbook.ActiveSheet
'    Set ActCell = sh.Application.ActiveCell
'    ' Снять защиту листа
'    If sh.ProtectContents = True Then
'        sh.Unprotect Pass(sh)
'    End If
'    ' Получить диапазоны функции EvDRE
'    GetRanges sh, EVDRE()
'    ' Обновить формулы
'    sh.Calculate
'    InProcess = True
'End Sub
'
'Private Sub SetNormalMode(sh, ActCell)
''
'' Снять основные параметры для оптимизации выполнения макросов
''
'    Application.EnableEvents = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.CutCopyMode = False
'    'Установить защиту листа
'    If sh.ProtectContents = False And sh.Columns("A:A").Hidden = True And Not Event_Name Like "BEFORE*" Then
'        sh.Protect Pass(sh), DrawingObjects:=True, Contents:=True, Scenarios:=True _
'            , AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
'    End If
'    ' Обновить формулы
'    sh.Calculate
'    ' Вернуть позицию курсора
'    sh.Activate
'    On Error Resume Next
'    sh.Cells(ActCell.Row, ActCell.Column).Select
'    On Error GoTo 0
'    ' Если событие AFTER
'    If Not Event_Name Like "BEFORE*" Then
'        Application.ScreenUpdating = True
'        InProcess = False
'    End If
'    ' Показать результат функций EVDRE в StatusBar
'    Do
'        i = i + 1
'    Loop Until i = UBound(EVDRE, 2) Or sh.Range(EVDRE(0, i)).Value <> "EVDRE:OK"
'    Application.StatusBar = Range(EVDRE(0, i)).Value
'End Sub
'
'Private Sub GetRanges(sh, EVDRE())
''
'' Получить диапазоны функции EvDRE
''
'    Dim cnt As Integer
'    cnt = 0
'    ReDim Preserve EVDRE(0 To 8, 0 To cnt)
'    EVDRE(0, cnt) = "EVDRE_Address"
'    EVDRE(1, cnt) = "AppName"
'    EVDRE(2, cnt) = "KeyRange"
'    EVDRE(3, cnt) = "ExpandRange"
'    EVDRE(4, cnt) = "PageKeyRange"
'    EVDRE(5, cnt) = "ColKeyRange"
'    EVDRE(6, cnt) = "RowKeyRange"
'    EVDRE(7, cnt) = "CellKeyRange"
'    'EVDRE(8, cnt) = "Data"
'
'    ' Поиск функций EvDRE на листе
'    Set e = Cells.Find("=EvDRE", Cells(sh.Rows.Count, sh.Columns.Count), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False)
'    If Not e Is Nothing And e.Formula <> "=EVDRE()" Then
'        firstAddress = e.Address
'        Do
'            If e.Text <> e.Formula Then
'                cnt = cnt + 1
'                ReDim Preserve EVDRE(0 To 8, 0 To cnt)
'                ' Нумерация функций EVDRE на листе
'                If IsEmpty(Cells(e.Row, e.Column + 1).Value) Or IsNumeric(Cells(e.Row, e.Column + 1).Value) Then
'                    Cells(e.Row, e.Column + 1).Value = cnt
'                End If
'
'                ' Получить параметры EVDRE
'                EVDRE(0, cnt) = e.Address       'EVDRE_Address
'                s = e.Formula
'                s = Replace(s, ")", "", 8)
'                a = Split(s, ",")
'                EVDRE(1, cnt) = a(0)            'AppName
'                EVDRE(2, cnt) = a(1)            'KeyRange
'                If UBound(a) = 2 Then
'                    EVDRE(3, cnt) = a(2)        'ExpandRange
'                End If
'
'                ' Получить параметры остальных диапазонов
'                For i = 4 To UBound(EVDRE, 1)
'                    Set f = sh.Range(a(1)).Find(EVDRE(i, 0), LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'                    s = sh.Cells(f.Row, f.Column + 1).Formula
'                    s = Replace(s, ")", "", 8)
'                    EVDRE(i, cnt) = Split(s, ",")
'                Next i
'            End If
'            Set e = Cells.Find("=EvDRE", e, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False)
'        Loop While firstAddress <> e.Address
'    End If
'End Sub
'
'Private Sub GetRangeBounds(RangeName As String, Nums As Variant, TopRow As Variant, LeftCol As Variant, BottomRow As Variant, RightCol As Variant)
''
'' Получить границы нескольких диапазонов
''
'    ' Обязательная (повторная) инициализация
'    TopRow = 99999
'    LeftCol = 99999
'    BottomRow = 0
'    RightCol = 0
'
'    ' Поиск номера диапазона в массиве EVDRE
'    For i = 5 To UBound(EVDRE, 1)
'        If RangeName = EVDRE(i, 0) Then RangeNum = i
'    Next i
'
'    ' Цикл по параметрам Nums
'    For i = 0 To UBound(Nums)
'        ' Определение номера EVDRE и номера поддиапазона
'        n = Split(Nums(i), ".")
'        If UBound(n) = 0 Then
'            EvdreNum = n(0)
'            MultiRangeNum = 0
'        ElseIf UBound(n) = 1 Then
'            EvdreNum = 1
'            If RangeName = "RowKeyRange" Then MultiRangeNum = n(0)
'            If RangeName = "ColKeyRange" Then MultiRangeNum = n(1)
'        Else
'            EvdreNum = n(0)
'            If RangeName = "RowKeyRange" Then MultiRangeNum = n(1)
'            If RangeName = "ColKeyRange" Then MultiRangeNum = n(2)
'        End If
'
'        ' Выбор конкретного диапазона
'        o = EVDRE(RangeNum, EvdreNum)
'        ' Выбор конкретного поддиапазона, если указан
'        If MultiRangeNum > 0 Then o = Array(o(MultiRangeNum - 1))
'
'        ' Цикл по поддиапазонам
'        For j = 0 To UBound(o)
'            ' Получить границы одного поддиапазона
'            a = Split(Replace(o(j), "$", ""), ":")
'            ' Если диапазон из одной ячейки
'            If UBound(a) = 0 Then a = Array(a(0), a(0))
'            ' Отобрать самую верхнюю строку
'            TopRow = WorksheetFunction.Min(TopRow, Range(a(0)).Row)
'            ' Отобрать самую левую колонку
'            LeftCol = WorksheetFunction.Min(LeftCol, Range(a(0)).Column)
'            ' Отобрать самую нижнюю строку
'            BottomRow = WorksheetFunction.Max(BottomRow, Range(a(1)).Row)
'            ' Отобрать самую правую колонку
'            RightCol = WorksheetFunction.Max(RightCol, Range(a(1)).Column)
'        Next j
'    Next i
'    ' Если диапазон пуст
'    If TopRow = 99999 Then TopRow = 0
'    If LeftCol = 99999 Then LeftCol = 0
'End Sub
'
'Sub DeleteRows(sh)
''
'' Удалить строки
''
'    ' Поиск столбца-фильтра
'    Set f = sh.Rows(1).Find("Delete", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Определить номера EVDRE для которых применить образец
'        Nums = Split(Replace(f.Value, " ", "", 7), ",")
'        If UBound(Nums) = -1 Then Nums = Array(1)
'        ' Получить границы дипазона ColKeyRange
'        GetRangeBounds "ColKeyRange", Nums, ckr_TopRow, ckr_LeftCol, ckr_BottomRow, ckr_RightCol
'        ' Получить границы дипазона RowKeyRange
'        GetRangeBounds "RowKeyRange", Nums, rkr_TopRow, rkr_LeftCol, rkr_BottomRow, rkr_RightCol
'        ' Получить границы дипазона CellKeyRange
'        GetRangeBounds "CellKeyRange", Nums, clkr_TopRow, clkr_LeftCol, clkr_BottomRow, clkr_RightCol
'        ' Расчет последней колонки с учетом CellKeyRange
'        If clkr_RightCol > 0 Then ckr_RightCol = ckr_RightCol + clkr_RightCol - clkr_LeftCol + 1
'
'        ' Обновить формулы
'        sh.Calculate
'        ' Диапазон который отобразить
'        sh.Cells.EntireRow.Hidden = False
'        ' Диапазон который удалить
'        Set r = sh.Columns(f.Column)
'        r.Hidden = False
'        With r
'            Set f = .Find(1, LookIn:=xlValues, LookAt:=xlWhole)
'            If Not f Is Nothing Then
'                firstAddress = f.Address
'                's = rkr_LeftCol & f.Row & ":" & ckr_RightCol & f.Row
'                Set ur = Range(sh.Cells(f.Row, rkr_LeftCol), sh.Cells(f.Row, ckr_RightCol))
'                Do
'                    Set f = .FindNext(f)
'                    's = rkr_LeftCol & f.Row & ":" & ckr_RightCol & f.Row
'                    'Set dr = Range(s)
'                    Set dr = Range(sh.Cells(f.Row, rkr_LeftCol), sh.Cells(f.Row, ckr_RightCol))
'                    Set ur = Union(ur, dr)
'                Loop Until f.Address = firstAddress
'                ur.Select
'                ' Удалить строки
'                ur.Delete Shift:=xlUp
'            End If
'        End With
'        r.Hidden = True
'    End If
'End Sub
'
'Private Sub HideRows(sh)
''
'' Скрыть строки
''
'    Dim r As Range
'    ' Поиск столбца-фильтра
'    Set f = sh.Rows(1).Find("Hide", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Обновить формулы
'        sh.Calculate
'        ' Диапазон который отобразить
'        sh.Cells.EntireRow.Hidden = False
'        ' Диапазон который скрыть
'        Set r = sh.Columns(f.Column)
'        r.Hidden = False
'        With r
'            Set f = .Find(1, LookIn:=xlValues, LookAt:=xlWhole)
'            If Not f Is Nothing Then
'                firstAddress = f.Address
'                Set ur = Rows(f.Row)
'                Do
'                    Set f = .FindNext(f)
'                    Set ur = Union(ur, Rows(f.Row))
'                Loop Until f.Address = firstAddress
'                ' Скрыть строки
'                ur.EntireRow.Hidden = True
'            End If
'        End With
'        r.Hidden = True
'    End If
'End Sub
'
'Private Sub HideColumns(sh, ColNum As Variant)
''
'' Скрыть столбцы
''
'    Dim r As Range
'    ' Поиск строки-фильтра
'    Set f = sh.Columns(ColNum).Find("HideColumns", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Обновить формулы
'        sh.Calculate
'        ' Диапазон который отобразить
'        'sh.Cells.EntireColumn.Hidden = False
'        Set uh = Columns(ColNum)
'        Set uh = uh.Resize(sh.Rows.Count, sh.Columns.Count - ColNum)
'        uh.EntireColumn.Hidden = False
'
'        ' Диапазон который скрыть
'        Set ur = Columns(ColNum)
'        Set r = sh.Rows(f.Row)
'        r.Hidden = False
'        With r
'            Set f = .Find(1, LookIn:=xlValues, LookAt:=xlWhole)
'            If Not f Is Nothing Then
'                firstAddress = f.Address
'                Set ur = Union(ur, Columns(f.Column))
'                Do
'                    Set f = .FindNext(f)
'                    Set ur = Union(ur, Columns(f.Column))
'                Loop Until f.Address = firstAddress
'                ' Скрыть столбцы
'                ur.EntireColumn.Hidden = True
'            End If
'        End With
'        r.Hidden = True
'    End If
'End Sub
'
'Private Sub MergeCells(sh, ColNum As Variant)
''
'' Объединить ячейки
''
'    ' Поиск строки-указателя
'    Set f = sh.Columns(ColNum).Find("Merge", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        Application.DisplayAlerts = False
'        fAddress = f.Address
'        ' Обработка одной строки
'        Do
'            Set f2 = sh.Rows(f.Row).Find("*", f, LookIn:=xlValues)
'            If Not f2 Is Nothing Then
'                Set prev = f2
'                firstAddress = f2.Address
'                ' Объединение ячеек в одной строке
'                Do
'                    If f2.Text = prev.Text Then
'                        Range(f2, prev).MergeCells = True
'                    Else
'                        Set prev = f2
'                    End If
'                    Set f2 = sh.Rows(f2.Row).Find("*", f2, LookIn:=xlValues)
'                Loop While firstAddress <> f2.Address
'                ' Объединение ячеек с верхней строкой
'                'If prevM = f.Row - 1 Then
'                '    Set f2 = sh.Rows(f.Row - 1).Find("*", f2, LookIn:=xlValues)
'                '    Do
'                '        Set foll = Cells(f2.Row + 1, f2.Column)
'                '        If f2.Text = foll.Text Or Len(foll.Text) = 0 Then
'                '            Range(f2, foll).MergeCells = True
'                '        End If
'                '        Set f2 = sh.Rows(f2.Row).Find("*", f2, LookIn:=xlValues)
'                '    Loop While firstAddress <> f2.Address
'                'End If
'            End If
'            'prevM = f.Row
'            ' Поиск следующей строки для объединения
'            Set f = sh.Columns(ColNum).Find("Merge", Cells(f.Row, ColNum), LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'        Loop While fAddress <> f.Address
'        Application.DisplayAlerts = True
'    End If
'End Sub
'
'Private Sub UnMergeCells(sh, ColNum As Variant)
''
'' Снять объединение ячеек
''
'    ' Поиск строки-указателя
'    Set f = sh.Columns(ColNum).Find("UnMerge", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        fAddress = f.Address
'        ' Обработка одной строки
'        Do
'            Rows(f.Row).MergeCells = False
'            ' Поиск следующей строки для объединения
'            Set f = sh.Columns(ColNum).Find("UnMerge", Cells(f.Row, ColNum), LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'        Loop While fAddress <> f.Address
'    End If
'End Sub
'
'Private Sub CopyPaste(sh, ColNum As Variant)
''
'' Скопировать содержимое в другую колонку
''
'    ' Поиск строки с CopyPaste
'    Set f = sh.Columns(ColNum).Find("CopyPaste", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'    If Not f Is Nothing Then
'        fAddress = f.Address
'        ' Обработка одной строки
'        Do
'            ' Определить номер EVDRE для которого применяется формула
'            a = Split(f.Value, " ")
'            Nums = Split(Replace(Replace(f.Value, a(0), ""), " ", ""), ",")
'            If UBound(Nums) = -1 Then Nums = Array("1")
'
'            ' Применить формулы для отдельных поддиапазонов
'            For i = 0 To UBound(Nums)
'                ' Получить границы дипазона ColKeyRange
'                GetRangeBounds "ColKeyRange", Array(Nums(i)), ckr_TopRow, ckr_LeftCol, ckr_BottomRow, ckr_RightCol
'                ' Получить границы дипазона RowKeyRange
'                GetRangeBounds "RowKeyRange", Array(Nums(i)), rkr_TopRow, rkr_LeftCol, rkr_BottomRow, rkr_RightCol
'
'                ' Скопировать
'                Set fc = sh.Rows(f.Row).Find("Copy", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'                If Not fc Is Nothing Then
'                    Range(sh.Cells(rkr_TopRow, fc.Column), sh.Cells(rkr_BottomRow, fc.Column)).Copy
'                End If
'
'                ' Вставить
'                Set fp = sh.Rows(f.Row).Find("Paste", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'                If Not fp Is Nothing Then
'                    Set dr = Range(sh.Cells(rkr_TopRow, fp.Column), sh.Cells(rkr_BottomRow, fp.Column))
'                    dr.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'                End If
'            Next i
'
'            ' Скрыть строку с CopyPaste
'            Rows(f.Row).EntireRow.Hidden = True
'            ' Поиск следующей строки с формулами
'            Set f = sh.Columns(ColNum).Find("CopyPaste", f, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'        Loop While fAddress <> f.Address
'    End If
'End Sub
'
'Private Sub ApplyFormulas(sh, ColNum As Variant)
''
'' Распространить формулы в колонках
''
'    ' Поиск строки с формулами
'    Set f = sh.Columns(ColNum).Find("Formula", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'    If Not f Is Nothing Then
'        fAddress = f.Address
'        ' Обработка одной строки
'        Do
'            ' Определить номер EVDRE для которого применяется формула
'            a = Split(f.Value, " ")
'            Nums = Split(Replace(Replace(f.Value, a(0), ""), " ", ""), ",")
'            If UBound(Nums) = -1 Then Nums = Array("1")
'
'            ' Применить формулы для отдельных поддипазонов
'            For i = 0 To UBound(Nums)
'                ' Получить границы дипазона ColKeyRange
'                GetRangeBounds "ColKeyRange", Array(Nums(i)), ckr_TopRow, ckr_LeftCol, ckr_BottomRow, ckr_RightCol
'                ' Получить границы дипазона RowKeyRange
'                GetRangeBounds "RowKeyRange", Array(Nums(i)), rkr_TopRow, rkr_LeftCol, rkr_BottomRow, rkr_RightCol
'                ' Получить границы дипазона RowKeyRange
'                GetRangeBounds "CellKeyRange", Array(Nums(i)), clkr_TopRow, clkr_LeftCol, clkr_BottomRow, clkr_RightCol
'                ' Если не указаны диапазоны применить формулы и к заголовкам строк
'                If Len(Nums(i)) = Len(Replace(Nums(i), ".", "")) Then ckr_LeftCol = rkr_LeftCol
'                ' Если не указаны диапазоны применить формулы и к CellKeyRange
'                If Len(Nums(i)) = Len(Replace(Nums(i), ".", "")) And clkr_RightCol > 0 Then:
'                    ckr_RightCol = ckr_RightCol + clkr_RightCol - clkr_LeftCol + 1
'                    'ckr_RightCol = ckr_RightCol + ckr_RightCol - ckr_LeftCol + 1
'                ' Скопировать формулы
'                sh.Range(sh.Cells(f.Row, ckr_LeftCol), sh.Cells(f.Row, ckr_RightCol)).Copy
'                'Вставить формулы
'                Set dr = Range(sh.Cells(rkr_TopRow, ckr_LeftCol), sh.Cells(rkr_BottomRow, ckr_RightCol))
'                dr.PasteSpecial Paste:=xlPasteFormulas, SkipBlanks:=True, Transpose:=False
'            Next i
'
'            ' Обновить формулы
'            sh.Calculate
'            ' Скрыть строку с формулой
'            Rows(f.Row).EntireRow.Hidden = True
'            ' Поиск следующей строки с формулами
'            Set f = sh.Columns(ColNum).Find("Formula", f, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'        Loop While fAddress <> f.Address
'    End If
'End Sub
'
'Sub Sort(sh)
''
'' Сортировать строки
''
'    ' Поиск столбца-фильтра
'    Set f = sh.Rows(1).Find("Sort", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Определить номера EVDRE для которых применить образец
'        Nums = Split(Replace(f.Value, " ", "", 5), ",")
'        If UBound(Nums) = -1 Then Nums = Array(1)
'        ' Получить границы дипазона ColKeyRange
'        GetRangeBounds "ColKeyRange", Nums, ckr_TopRow, ckr_LeftCol, ckr_BottomRow, ckr_RightCol
'        ' Получить границы дипазона RowKeyRange
'        GetRangeBounds "RowKeyRange", Nums, rkr_TopRow, rkr_LeftCol, rkr_BottomRow, rkr_RightCol
'        ' Получить границы дипазона CellKeyRange
'        GetRangeBounds "CellKeyRange", Nums, clkr_TopRow, clkr_LeftCol, clkr_BottomRow, clkr_RightCol
'        ' Расчет последней колонки с учетом CellKeyRange
'        If clkr_RightCol > 0 Then ckr_RightCol = ckr_RightCol + clkr_RightCol - clkr_LeftCol + 1
'
'        ' Диапазон который отобразить
'        sh.Cells.EntireRow.Hidden = False
'        ' Очистить существующие фильтры
'        sh.Sort.SortFields.clear
'        ' Колонка/ключ сортировки
'        Set akr = Range(sh.Cells(rkr_TopRow, f.Column), sh.Cells(rkr_BottomRow, f.Column))
'        sh.Sort.SortFields.Add Key:=akr, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'        ' Диапазон сортировки
'        Set sr = Range(sh.Cells(rkr_TopRow, rkr_LeftCol), sh.Cells(rkr_BottomRow, ckr_RightCol))
'        With sh.Sort
'            .SetRange sr
'            .Header = xlGuess
'            .MatchCase = False
'            .Orientation = xlTopToBottom
'            .SortMethod = xlPinYin
'            .Apply
'        End With
'    End If
'End Sub
'
'Private Sub Freeze_Unfreeze()
''
'' Закрепить или открепить область
''
'    ' Снять защиту листа
'    If ActiveSheet.ProtectContents = True Then
'        ActiveSheet.Unprotect Pass(ActiveSheet)
'    End If
'    ' Закрепить или открепить область
'    If ActiveWindow.FreezePanes Then
'        ActiveWindow.FreezePanes = False
'    Else
'        ActiveWindow.FreezePanes = True
'    End If
'    'Установить защиту листа
'    If ActiveSheet.ProtectContents = False And ActiveSheet.Columns("A:A").Hidden = True Then
'        ActiveSheet.Protect Pass(ActiveSheet), DrawingObjects:=True, Contents:=True, Scenarios:=True _
'            , AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
'    End If
'End Sub
'
'Private Sub AutoFilter()
''
'' Включить или выключить автофильтр
''
'    ' Снять защиту листа
'    If ActiveSheet.ProtectContents = True Then
'        ActiveSheet.Unprotect Pass(ActiveSheet)
'    End If
'    ' Включить или выключить автофильтр
'    On Error Resume Next
'    Selection.AutoFilter
'    If Err.Number Then MsgBox "Выделите диапазон заголовка фильтра", vbCritical, "Автофильтр"
'    On Error GoTo 0
'    'Установить защиту листа
'    If ActiveSheet.ProtectContents = False And ActiveSheet.Columns("A:A").Hidden = True Then
'        ActiveSheet.Protect Pass(ActiveSheet), DrawingObjects:=True, Contents:=True, Scenarios:=True _
'            , AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
'    End If
'End Sub
'
'Private Sub ShowAll()
''
'' Отобразить строки и столбцы
''
'    Application.Calculation = xlCalculationManual
'    Set sh = ActiveWorkbook.ActiveSheet
'    If sh.ProtectContents = True Then
'        sh.Unprotect
'    End If
'    sh.Rows.Hidden = False
'    sh.Columns.Hidden = False
'    Application.Calculation = xlCalculationAutomatic
'End Sub
'
'Private Sub ProtectBook()
''
'' Установить пароль на все листы рабочей книги
''
'    Dim argument As String
'    ' Цикл по всем листам книги
'    For i = 1 To ActiveWorkbook.Sheets.Count
'        If Sheets(i).Visible = xlSheetVisible Then
'            Sheets(i).Activate
'            t = AFTER_EXPAND(argument)
'        End If
'    Next i
'End Sub
'
'Private Sub UnProtectBook()
''
'' Снять пароль со всех листов рабочей книги
''
'    Event_Name = "BEFORE_"
'    ' Установить основные параметры для оптимизации выполнения макросов
'    SetOptimizeMode sh, ActCell
'    ' Цикл по всем листам книги
'    For i = 1 To ActiveWorkbook.Sheets.Count
'        If Sheets(i).Visible = True And Sheets(i).ProtectContents = True Then
'            Sheets(i).Unprotect ActCell.Value
'        End If
'    Next i
'    sh.Activate
'    ' Снять основные параметры для оптимизации выполнения макросов
'    SetNormalMode sh, ActCell
'End Sub
'
'Private Sub NewRows(nt As Variant)
''
'' Добавить новые строки
''
'    Set sh = ActiveWorkbook.ActiveSheet
'
'    ' Поиск основной колонки NewRow
'    Set f = sh.Rows(1).Find("NewRows", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'    If f Is Nothing Then
'        MsgBox "Не найдена колонка NewRows", vbCritical, "Добавление строки"
'    Else
'        ' Определить количество вставляемых строк
'        q = Replace(f.Value, " ", "", 8)
'        If IsNumeric(q) = False Then q = 5
'        ' Поиск строки-образца
'        Set t = sh.Columns(f.Column).Find(nt & "Template", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'        If t Is Nothing Then
'            MsgBox "Не найдена строка-образец", vbCritical, "Добавление строки"
'        Else
'            ' Установить основные параметры для оптимизации выполнения макросов
'            SetOptimizeMode sh, ActCell
'
'            ' Определить номера EVDRE для которых применить образец
'            Nums = Split(Replace(t.Value, " ", "", Len(nt) + 9), ",")
'            If UBound(Nums) = -1 Then Nums = Array(nt)
'            ' Получить границы дипазона ColKeyRange
'            GetRangeBounds "ColKeyRange", Nums, ckr_TopRow, ckr_LeftCol, ckr_BottomRow, ckr_RightCol
'            ' Получить границы дипазона RowKeyRange
'            GetRangeBounds "RowKeyRange", Nums, rkr_TopRow, rkr_LeftCol, rkr_BottomRow, rkr_RightCol
'            ' Получить границы дипазона CellKeyRange
'            GetRangeBounds "CellKeyRange", Nums, clkr_TopRow, clkr_LeftCol, clkr_BottomRow, clkr_RightCol
'            ' Расчет последней колонки с учетом CellKeyRange
'            If clkr_RightCol > 0 Then ckr_RightCol = ckr_RightCol + clkr_RightCol - clkr_LeftCol + 1
'
'            'Поиск последней строки без AfterRange
'            Do
'                Set lr = Range(sh.Cells(rkr_BottomRow, rkr_LeftCol), sh.Cells(rkr_BottomRow, rkr_RightCol))
'                Set l = lr.Find("EV_AFTER", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'                If Not l Is Nothing Then
'                    rkr_BottomRow = rkr_BottomRow - 1
'                End If
'            Loop While Not l Is Nothing And rkr_BottomRow >= rkr_TopRow
'
'            ' Сформировать адрес последней строки
'            Set lr = Range(sh.Cells(rkr_BottomRow, rkr_LeftCol), sh.Cells(rkr_BottomRow, ckr_RightCol))
'            ' Вставить необходимое количество строк
'            For i = 1 To q
'                lr.Insert Shift:=xlDown
'            Next
'            ' Скопировать последнюю строку
'            lr.Copy
'            ' Сформировать адрес последней строки
'            Set lr = Range(sh.Cells(rkr_BottomRow, rkr_LeftCol), sh.Cells(rkr_BottomRow, ckr_RightCol))
'            ' Вставить все из последней строки
'            lr.PasteSpecial Paste:=xlPasteAll, SkipBlanks:=False, Transpose:=False
'
'            ' Сформировать адрес вставленных строк
'            Set tr = Range(sh.Cells(rkr_BottomRow + 1, rkr_LeftCol), sh.Cells(rkr_BottomRow + q, ckr_RightCol))
'            ' Скопировать строку-образец
'            Range(sh.Cells(t.Row, rkr_LeftCol), sh.Cells(t.Row, ckr_RightCol)).Copy
'            ' Вставить все из строки-образец
'            tr.PasteSpecial Paste:=xlPasteAll, SkipBlanks:=False, Transpose:=False
'
'            ' Снять основные параметры для оптимизации выполнения макросов
'            SetNormalMode sh, ActCell
'        End If
'    End If
'End Sub
'
'Private Sub AddNewRows1()
'    NewRows 1
'End Sub
'
'Private Sub AddNewRows2()
'    NewRows 2
'End Sub
'
'Private Sub AddNewRows3()
'    NewRows 3
'End Sub
'
'Private Sub AddNewRows4()
'    NewRows 4
'End Sub
'
'Private Sub AddNewRows5()
'    NewRows 5
'End Sub
'
'Private Sub AddNewRows6()
'    NewRows 6
'End Sub
'
'Private Sub AddNewRows7()
'    NewRows 7
'End Sub
'
'Private Sub AddNewRows8()
'    NewRows 8
'End Sub
'
'Private Sub AddNewRows9()
'    NewRows 9
'End Sub
'
'Private Sub ChangeRow(nt As Variant)
''
'' Изменить существующую строку
''
'    Set sh = ActiveWorkbook.ActiveSheet
'
'    ' Поиск основной колонки NewRow
'    Set f = sh.Rows(1).Find("NewRows", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'    If f Is Nothing Then
'        MsgBox "Не найдена колонка NewRows", vbCritical, "Добавление строки"
'    Else
'        ' Поиск строки-образца
'        Set t = sh.Columns(f.Column).Find(nt & "Change", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'        If t Is Nothing Then
'            MsgBox "Не найдена строка-образец", vbCritical, "Добавление строки"
'        Else
'            ' Установить основные параметры для оптимизации выполнения макросов
'            SetOptimizeMode sh, ActCell
'
'            ' Определить номера EVDRE для которых применить образец
'            Nums = Split(Replace(t.Value, " ", "", Len(nt) + 7), ",")
'            If UBound(Nums) = -1 Then Nums = Array(nt)
'            ' Получить границы дипазона ColKeyRange
'            GetRangeBounds "ColKeyRange", Nums, ckr_TopRow, ckr_LeftCol, ckr_BottomRow, ckr_RightCol
'            ' Получить границы дипазона RowKeyRange
'            GetRangeBounds "RowKeyRange", Nums, rkr_TopRow, rkr_LeftCol, rkr_BottomRow, rkr_RightCol
'            ' Получить границы дипазона CellKeyRange
'            GetRangeBounds "CellKeyRange", Nums, clkr_TopRow, clkr_LeftCol, clkr_BottomRow, clkr_RightCol
'            ' Расчет последней колонки с учетом CellKeyRange
'            If clkr_RightCol > 0 Then ckr_RightCol = ckr_RightCol + clkr_RightCol - clkr_LeftCol + 1
'
'            ' Сформировать адрес вставляемой строки
'            If rkr_BottomRow = ActCell.Row Then r = ActCell.Row Else r = ActCell.Row + 1
'            Set ir = Range(sh.Cells(r, rkr_LeftCol), sh.Cells(r, ckr_RightCol))
'            ' Вставить новую строку
'            ir.Insert Shift:=xlDown
'            ' Сформировать адрес изменяемой строки
'            Set cr = Range(sh.Cells(ActCell.Row, rkr_LeftCol), sh.Cells(ActCell.Row, ckr_RightCol))
'            ' Скопировать изменяемую строку
'            cr.Copy
'            ' Сформировать адрес вставляемой строки
'            Set ir = Range(sh.Cells(r, rkr_LeftCol), sh.Cells(r, ckr_RightCol))
'            ' Вставить все из изменяемой строки
'            ir.PasteSpecial Paste:=xlPasteAll, SkipBlanks:=False, Transpose:=False
'
'            ' Скопировать строку-образец
'            Range(sh.Cells(t.Row, rkr_LeftCol), sh.Cells(t.Row, ckr_RightCol)).Copy
'            ' Сформировать адрес новой строки
'            'Set tr = Range(sh.Cells(ActCell.Row, rkr_LeftCol), sh.Cells(ActCell.Row, ckr_RightCol))
'            If rkr_BottomRow = ActCell.Row - 1 Then r = ActCell.Row
'            Set ir = Range(sh.Cells(r, rkr_LeftCol), sh.Cells(r, ckr_RightCol))
'            ' Вставить все из строки-образец
'            ir.PasteSpecial Paste:=xlPasteAll, SkipBlanks:=True, Transpose:=False
'
'            ' Скрыть строки
'            HideRows sh
'            Application.EnableEvents = True
'
'            ' Найти значимые колонки
'            Set ckr = Range(sh.Cells(ckr_BottomRow, ckr_LeftCol), sh.Cells(ckr_BottomRow, ckr_RightCol))
'            Set v = ckr.Find("*", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'            If Not v Is Nothing Then
'                firstAddress = v.Address
'                Set cc = Cells(r - 1, v.Column)
'                Set ur = cc
'                Do
'                    'v.Select
'                    Set cc = Cells(r - 1, v.Column)
'                    ' Обнулить изменяемую строку
'                    If Not cc.Formula Like "=*" Then Set ur = Union(ur, cc) 'cc.ClearContents
'                    Set v = ckr.Find("*", v, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'                Loop While firstAddress <> v.Address
'                ur.ClearContents
'            End If
'
'            ' Снять основные параметры для оптимизации выполнения макросов
'            SetNormalMode sh, ActCell
'        End If
'    End If
'End Sub
'
'Private Sub ChangeRow1()
'    ChangeRow 1
'End Sub
'
'Private Sub ChangeRow2()
'    ChangeRow 2
'End Sub
'
'Private Sub ChangeRow3()
'    ChangeRow 3
'End Sub
'
'Private Sub ChangeRow4()
'    ChangeRow 4
'End Sub
'
'Private Sub ChangeRow5()
'    ChangeRow 5
'End Sub
'
'Private Sub ChangeRow6()
'    ChangeRow 6
'End Sub
'
'Private Sub ChangeRow7()
'    ChangeRow 7
'End Sub
'
'Private Sub ChangeRow8()
'    ChangeRow 8
'End Sub
'
'Private Sub ChangeRow9()
'    ChangeRow 9
'End Sub
'
'Private Function SAPlogon()
''
'' Вход в систему SAP
''
'    Application.StatusBar = "Соединение с сервером..."
'    SAPlogon = False
'    On Error Resume Next
'    Workbooks.Open "C:\PROGRAM FILES\COMMON FILES\SAP SHARED\BW\sapbex.xla"
'    If Err.Number Then MsgBox "Не найден файл sapbex.xla", vbCritical, "Подключение к серверу SAP": Exit Function
'    On Error GoTo 0
'    Set myConnection = Run("SAPBEX.XLA!SAPBEXgetConnection")
'    With myConnection
'        Select Case Application.Run("EVSVR")
'            Case "HTTP://172.16.34.1"
'                SapServer = "172.16.10.4"
'                .Client = "100"
'            Case "HTTP://172.16.10.12"
'                SapServer = "172.16.10.12"
'                .Client = "200"
'            Case "HTTP://172.16.30.6"
'                SapServer = "172.16.30.6"
'                .Client = "300"
'            Case "HTTP://V-SAP-DBI"
'                SapServer = "172.16.10.4"
'                .Client = "100"
'            Case "HTTP://V-SAP-QBI"
'                SapServer = "172.16.10.12"
'                .Client = "200"
'            Case "HTTP://V-SAP-PBI"
'                SapServer = "172.16.30.6"
'                .Client = "300"
'        End Select
'        .ApplicationServer = SapServer
'        .SystemNumber = "00"
'        .User = "WF-COMM"
'        .Password = "P@ssw0rd"
'        .Language = "en"
'        .logon 0, True
'        If .IsConnected <> 1 Then
'            .logon 0, False
'            If .IsConnected <> 1 Then MsgBox "Соединение с системой установить не удалось", vbCritical, "Подключение к серверу SAP": Exit Function
'        End If
'    End With
'    SAPlogon = True
'    Application.StatusBar = "Соединение с сервером успешно установлено"
'End Function
'
'Private Function SAPconnect(fm_name)
''
'' Подключение к системе SAP, создание объекта Функциональный модуль
''
'    'SAPconnect = False
'    If SAPlogon Then
'        Set sap = CreateObject("SAP.Functions")
'        sap.Connection = myConnection
'        On Error Resume Next '?
'        Set fm = sap.Add(fm_name)
'        If Err.Number Then MsgBox "Функциональный модуль " & fm_name & " не найден", vbCritical, "Подключение к серверу SAP": Exit Function
'        On Error GoTo 0 '?
'        SAPconnect = True
'    End If
'End Function
'
'Private Sub Locks(ds, mode)
''
'' Основной Sub: зарегистрировать (+), отрегистрировать (-/--/---), получить список (?)
''
'    Const fm_name = "ZUJW_LOCK"
'    On Error Resume Next
'        If sap Is Empty Then SAPconnect (fm_name)
'        fm.exports("DATASRC") = ds
'        fm.exports("USER_ID") = UCase(Application.Run("EVUSR"))
'        fm.exports("ACTION") = mode
'        Set tabl = fm.Tables("TAB")
'        fm.call
'    On Error GoTo 0
'End Sub
'
'Public Sub Locks_On(sh, Target)
''
'' Зарегистрировать пользователя на форме ввода
''
'    If Not InProcess And Not IsRegistered And UsingLocks Then
'        Set sh = ActiveWorkbook.ActiveSheet
'        Set ds = Locks_DS(sh)
'        If Not ds Is Nothing Then
'            ' Установить основные параметры для оптимизации выполнения макросов
'            SetOptimizeMode sh, ActCell
'            Locks ds, "+"
'            Locks_ListInCell sh
'            Locks_IsRegistered sh, "Y"
'            ' Снять основные параметры для оптимизации выполнения макросов
'            SetNormalMode sh, ActCell
'        End If
'    End If
'End Sub
'
'Public Sub Locks_Off(sh As Worksheet, Target)
''
'' Отрегистрировать пользователя от все форм ввода в книге
''
'    If UsingLocks Then
'
'        ' Установить основные параметры для оптимизации выполнения макросов
'        SetOptimizeMode sh, ActCell
'
'        Set sht = sh
'        For Each sh In Sheets
'            If UsingLocks And sh.Visible = xlSheetVisible And Locks_IsRegistered(sh, "?") Then
'                Locks sh, "-"
'            End If
'        Next
'        Set sh = sht
'
'        ' Снять основные параметры для оптимизации выполнения макросов
'        SetNormalMode sh, ActCell
'
'    End If
'
'End Sub
'
'Private Sub Locks_OffBook()
''
'' Отрегистрировать пользователя от все форм ввода в книге
''
'    'Set sh = ActiveWorkbook.ActiveSheet
'
'    Set sht = sh
'    For Each sh In Sheets
'         If sh.Visible = xlSheetVisible And UsingLocks Then
'            Locks_OffSheet sh
'         End If
'    Next
'    Set sh = sht
'End Sub
'
'Private Sub Locks_OffSheet(sh)
''
'' Отрегистрировать пользователя от формы ввода
''
'If UsingLocks Then
'    ' Снять защиту листа
'    If sh.ProtectContents = True Then
'        sh.Unprotect Pass(sh)
'    End If
'    ' Если пользователь зарегистрирован
'    If Locks_IsRegistered(sh, "?") Then
'        Locks Locks_DS(sh), "-"
'        Locks_IsRegistered sh, "N"
'    End If
'    ' Вывести список зарегистрированных пользователей в ячейке
'    Locks_ListInCell sh
'    ' Установить защиту листа
'    If sh.ProtectContents = False And sh.Columns("A:A").Hidden = True And Not Event_Name Like "BEFORE*" Then
'        sh.Protect Pass(sh), DrawingObjects:=True, Contents:=True, Scenarios:=True _
'            , AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
'    End If
'End If
'
'End Sub
'
'Private Sub Locks_OffAll()
''
'' Отрегистрировать пользователя абсолютно от всех форм ввода
''
'    Locks sh, "--"
'End Sub
'
'Public Function Locks_ListInWindow(sh, Target)
''
'' Вывести список зарегистрированных пользователей в окне
''
'    If UsingLocks Then
'        InProcess = True
'        If Target = Locks_Result(sh) Then
'            Locks sh, "?"
'            'Собрать список пользователей для вывода
'            s = ""
'            For i = 1 To tabl.RowCount
'              s = s & vbCrLf & i & ". " & tabl(i, 2) & " - " & tabl(i, 3) & " " & tabl(i, 4)
'            Next i
'            ' Вывести список пользователей в окне
'            If i = 1 Then
'                MsgBox "Сейчас с формой никто не работает", vbInformation, "Одновременная работа"
'            Else
'                MsgBox "Сейчас с формой " & tabl(1, 1) & " работают пользователи:" & vbCrLf & s, vbExclamation, "Одновременная работа"
'            End If
'            Locks_ListInWindow = True
'        Else
'            Locks_ListInWindow = False
'        End If
'        InProcess = False
'    End If
'End Function
'
'Private Sub Locks_ListInCell(sh)
''
'' Вывести список зарегистрированных пользователей в ячейке
''
'    Set ds = Locks_DS(sh)
'    If Not ds Is Nothing Then
'        Locks ds, "?"
'        'Собрать список пользователей в строку
'        s = ""
'        For i = 1 To tabl.RowCount
'            s = s & ", " & tabl(i, 2)
'        Next i
'        ' Вывести список пользователей в ячейку
'        Set lr = Locks_Result(sh)
'        If i = 1 Then
'            lr.Value = "Сейчас с формой никто не работает"
'            With lr.Font
'                .Name = "Calibri"
'                .Size = 12
'                .Color = -16751616
'            End With
'        Else
'            lr.Value = "Сейчас с формой работают: " & Mid(s, 3)
'            With lr.Font
'                .Name = "Calibri"
'                .Size = 12
'                .Color = -16233056
'            End With
'        End If
'    End If
'End Sub
'
'Private Function Locks_Result(sh) As Object
''
'' Возвращает ячейку где список пользователей
''
'    Set f = sh.Rows(1).Find("Сейчас с формой", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
'    If f Is Nothing Then
'        Set f = sh.Rows(1).Find("", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
'    End If
'    Set Locks_Result = f
'End Function
'
'Public Function Locks_IsRegistered(sh, NewValue As String)
''
'' Синхронизация необходимости регистрации пользователя в ячейке и переменной
''
'    Set ds = Locks_DS(sh)
'    If Not ds Is Nothing Then
'        If NewValue = "Y" Then
'            Cells(ds.Row + 1, ds.Column).Value = "Y"
'            IsRegistered = True
'        ElseIf NewValue = "N" Then
'            sh.Cells(ds.Row + 1, ds.Column).Value = "N"
'            IsRegistered = False
'        ElseIf NewValue = "?" Then
'            If sh.Cells(ds.Row + 1, ds.Column).Value = "Y" Then
'                IsRegistered = True
'            Else
'                IsRegistered = False
'            End If
'            'Set sh = ActiveWorkbook.ActiveSheet '??
'        End If
'        Locks_IsRegistered = IsRegistered
'    Else
'        IsRegistered = True
'        Locks_IsRegistered = False
'    End If
'End Function
'
'Private Function Locks_DS(sh) As Object
''
'' Возвращает ячейку где код формы ввода (источник данных)
''
'    Set Locks_DS = sh.Rows(1).Find("DS_", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
'End Function
'
'Private Sub SetWorkStatus()
''
'' Установить рабочие статусы
''
'    Const fm_name = "ZUJW_STATUS"
'    Dim WorkStatus() As String
'    ReDim Preserve WorkStatus(0 To 9, -1 To 0)
'
'    ' Поиск колонки со статусами
'    Set f = Rows(1).Find("WorkStatus", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
'    If Not f Is Nothing Then
'        ' Установить основные параметры для оптимизации выполнения макросов
'        SetOptimizeMode sh, ActCell
'
'        ' Определить номера EVDRE
'        a = Split(f.Value, " ")
'        Nums = Split(Replace(Replace(f.Value, a(0), ""), " ", ""), ",")
'        If UBound(Nums) = -1 Then Nums = Array("1")
'
'        ' Собрать статусы с EVDRE
'        For i = 0 To UBound(Nums)
'            ' Получить границы дипазона ColKeyRange
'            GetRangeBounds "ColKeyRange", Nums, ckr_TopRow, ckr_LeftCol, ckr_BottomRow, ckr_RightCol
'            ' Получить границы дипазона RowKeyRange
'            GetRangeBounds "RowKeyRange", Array(Nums(i)), rkr_TopRow, rkr_LeftCol, rkr_BottomRow, rkr_RightCol
'            ' Инициализация массива статусов
'            For j = rkr_LeftCol To ckr_LeftCol - 1
'                v = Cells(rkr_TopRow - 1, j).Value
'                If Len(v) > 0 And v <> "1" And v <> "IncludeChildren" Then
'                    K = K + 1
'                    WorkStatus(K, -1) = j
'                    WorkStatus(K, 0) = Cells(rkr_TopRow - 1, j).Value
'                ElseIf v = "IncludeChildren" Then
'                    WorkStatus(6, -1) = j
'                    WorkStatus(6, 0) = "IncludeChildren"
'                End If
'            Next j
'            WorkStatus(7, 0) = "NewStatus"
'            WorkStatus(8, 0) = "ResultID"
'            WorkStatus(9, 0) = "ResultText"
'
'            ' Поиск рабочего статуса для изменения
'            Set r = Range(sh.Cells(rkr_TopRow, f.Column), sh.Cells(rkr_BottomRow, f.Column))
'            Set ws = r.Find("*", LookIn:=xlValues)
'            If Not ws Is Nothing Then
'                firstAddress = ws.Address
'                Do
'                    cnt = cnt + 1
'                    ' Добавление записи в массив статусов
'                    ReDim Preserve WorkStatus(0 To 9, -1 To cnt)
'                    ' Номер строки
'                    WorkStatus(0, cnt) = ws.Row
'                    ' Значение устанавливаемого статуса
'                    WorkStatus(7, cnt) = CStr(sh.Cells(ws.Row, f.Column).Value)
'                    ' Заполнение ключей
'                    For j = 1 To 6
'                        If Len(WorkStatus(j, -1)) <> 0 Then
'                            WorkStatus(j, cnt) = sh.Cells(ws.Row, CInt(WorkStatus(j, -1))).Value
'                        End If
'                    Next j
'                    Set ws = r.Find("*", ws, LookIn:=xlValues)
'                Loop Until ws.Address = firstAddress
'            End If
'
'            ' Вызов ФМ, который установит статусы из массива WorkStatus
'            ' Подключение к системе SAP, создание объекта Функциональный модуль
'            If SAPconnect(fm_name) Then
'                ' Инициализация параметров
'                Domain = Environ("UserDomain")
'                Select Case Domain
'                  Case "DTEK", "DTEKGROUP", "PAVLOGRADYGOL", "PES", "PRMZ", "SICHEV"
'                  Case Else: Domain = "DTEKGROUP"
'                End Select
'                fm.exports("USER_ID") = Domain & "\" & UCase(Application.Run("EVUSR"))
'                fm.exports("APPSET_ID") = Application.Run("EVAST")
'                fm.exports("APPLICATION_ID") = Range(EVDRE(1, Nums(i))).Value
'
'                ' Транспонирование массива статусов
'                Set st = fm.Tables("TAB")
'                For ii = 0 To cnt
'                  st.AppendRow
'                  For jj = 1 To 9
'                    st(ii + 1, jj) = WorkStatus(jj, ii)
'                  Next jj
'                Next ii
'
'                ' Вызов модуля в системе SAP
'                fm.call
'
'                ' Вывод результатов
'                jj = ckr_RightCol + 1
'                Range(sh.Cells(rkr_TopRow, jj), sh.Cells(rkr_BottomRow, jj + 1)).ClearContents
'                For ii = 1 To cnt
'                  temp = st(ii + 1, 8)
'                  sh.Cells(WorkStatus(0, ii), jj + 0) = st(ii + 1, 8)
'                  sh.Cells(WorkStatus(0, ii), jj + 1) = st(ii + 1, 9)
'                Next ii
'                Application.StatusBar = "SetWorkStatus - Ok:" & fm.imports("NUMOK") & ", Error:" & fm.imports("NUMERR")
'
'                ' Сохранить (комментарии) перед обновлением
'                Application.Run "MNU_eSUBMIT_REFSCHEDULE_SHEET_NOACTION"
'                ' Обновить форму после установки статусов
'                Application.Run "MNU_eTOOLS_EXPAND"
'                ' Установить основные параметры для оптимизации выполнения макросов
'                SetOptimizeMode sh, ActCell
'            End If
'        Next i
'
'        ' Снять основные параметры для оптимизации выполнения макросов
'        SetNormalMode sh, ActCell
'    End If
'End Sub
'
'Public Function MAXIF(SearchRange As Range, Criteria, MaxRange As Range)
''
'' Максимальное значение в ячейках удовлетворяющих условию
''
'' SearchRange - где ищется ключ
'' Criteria - что ищется в ключах
'' MaxRange - где ищется максимум
'
'    Application.Calculation = xlCalculationManual
'        Application.Volatile True 'строка которую сказали добавить Коним
'    MAXIF = 0
'
'    ' Диапазон в котором ищется ключ
'    Set c1 = ActiveWorkbook.ActiveSheet.Cells(SearchRange.Row, SearchRange.Column)
'    Set c2 = ActiveWorkbook.ActiveSheet.Cells(SearchRange.Row + SearchRange.Count - 1, SearchRange.Column)
'    Set r = ActiveWorkbook.ActiveSheet.Range(c1, c2)
'    With r
'        Set f = .Find(Criteria, LookIn:=xlValues, LookAt:=xlWhole)
'        If Not f Is Nothing Then
'            firstAddress = f.Address
'            Do
'                v = Cells(f.Row, MaxRange.Column).Value2
'                If MAXIF < v Then MAXIF = v
'                Set f = .Find(Criteria, f, LookIn:=xlValues, LookAt:=xlWhole)
'                'Set f = .FindNext(f)
'            Loop Until f.Address = firstAddress
'        End If
'    End With
'
'    Application.Calculation = xlCalculationAutomatic
'End Function
'
