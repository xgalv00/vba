VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''Private Sub Workbook_Activate()
''    Call onWorkBookActivate
'
''End Sub
'
''
'' Для унификации решения с макросами и простоты поддержки
''
'Private Sub Workbook_Open()
'
'' Системное событие - сразу после открытия книги
'' Загрузить с сервера набор макросов
'  'Dim c As Object
'  'For Each c In VBProject.VBComponents
'  '  If Left(c.Name, 9) = "MasterBPC" Then VBProject.VBComponents.Remove c
'  'Next c
'  'VBProject.VBComponents.Import "\\v-sap-pcts\Communic\Macros\MasterBPC.bas"
'End Sub
'
''
'' Для отслеживания одновременной работы пользователей с одними формами
''
'Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
'' Системное событие - после изменения ячейки на листе
'' Зарегистрировать пользователя на форме ввода
'    Locks_On sh, Target
'End Sub
'
Private Sub Workbook_BeforeClose(Cancel As Boolean)
'' Системное событие - перед закрытием книги
'' Отрегистрировать пользователя от всех форм в книге
'    Locks_Off sh, Target

    Call sendComments
End Sub
'
'Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
'' Системное событие - перед двойным нажатием мыши
'' Вывести список зарегистрированных пользователей в окне
'    Cancel = Locks_ListInWindow(sh, Target)
'End Sub
'
'Private Sub Workbook_SheetActivate(ByVal sh As Object)
'' Системное событие - после активации (другого) листа
'' Получить актуальное значение переменной IsRegistered
'    Locks_IsRegistered sh, "?"
'End Sub
'
''
'' Место вставки пользовательских макросов специфичных для конкретной формы
''
'Public Sub UserSub_Before(sh, ActCell, Event_Name, EVDRE())
'' Отрабатывает перед основным набором макросов
'
'End Sub
'
'Public Sub UserSub_After(sh, ActCell, Event_Name, EVDRE())
'' Отрабатывает после основного набора макросов
'End Sub
'
