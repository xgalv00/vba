Attribute VB_Name = "updateStatusModule"
Sub updateStatusFromChangeJournal()
    '
    Dim wbChange As Workbook, wbWork As Workbook
    Dim wsChange As Worksheet, wsWork As Worksheet
    '
    Set wbWork = ActiveWorkbook
    Set wsWork = wbWork.Sheets("Разработки")
    Set wbChange = Workbooks.Open("https://workspaces.dtek.com/it/oisup/ProjectSAP/ChangeManagement/Журнал регистрации изменений в проектах SAP.xlsm", , True)
    Set wsChange = wbChange.Sheets("журнал запросов на измение")
    '
    rowsCountChange = wsChange.UsedRange.Rows.Count + wsChange.UsedRange.Row
    rowsCountWork = wsWork.UsedRange.Rows.Count + wsWork.UsedRange.Row
    '
    countfix = 0
    '
    For curRowWork = 1 To rowsCountWork
        idTask = wsWork.Cells(curRowWork, 2)
        If idTask <> "" Then
            curStatusInWork = wsWork.Cells(curRowWork, 10)
            If curStatusInWork <> "6. Завершено" And curStatusInWork <> "7. Отменено" Then
                flagIsEnd = False
                For curRowChange = 1 To rowsCountChange
                    If wsChange.Cells(curRowChange, 2) = idTask Then
                        If wsChange.Cells(curRowChange, 15) = "Реализовано" Then
                            flagIsEnd = True
                            
                        End If
                    End If
                Next curRowChange
                If flagIsEnd = True Then
                    wsWork.Cells(curRowWork, 10) = "6. Завершено"
                    wsWork.Cells(curRowWork, 10).Interior.Color = 5296274
                    countfix = countfix + 1
                End If
            End If
        End If
    Next curRowWork
    '
    wbChange.Close False
    MsgBox "Изменено " & countfix & " задач."
    '
End Sub
