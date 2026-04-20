Attribute VB_Name = "mod_2_Constructor"
Option Explicit

' =========================================================
' БЛОК 1: ПОДГОТОВКА ДАННЫХ (Теперь по ИМЕНАМ)
' =========================================================
Function PrepareOrder() As clsOrder
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim n As New clsOrder
    With n
        ' Читаем данные по именам ячеек!
        .ID = wsP.Range("PZ_OrderNum").Text & wsP.Range("PZ_OrderPref").Text
        .Department = Trim(wsP.Range("PZ_Dept").Text)
        .WorkType = wsP.Range("PZ_WorkType").Text
        .ExtraInfo = wsP.Range("PZ_Extra").Text
    End With
    
    If n.ID = "" Or n.Department = "" Then
        MsgBox "Заполните № заказа и Цех!", 48: Set PrepareOrder = Nothing
    Else
        n.LoadFromRegistry
        If n.ItemCode = "" Then
            MsgBox "Артикул не найден!", 16: Set PrepareOrder = Nothing
        Else
            Set PrepareOrder = n
        End If
    End If
End Function

' =========================================================
' БЛОК 2: КНОПКИ СОЗДАНИЯ
' =========================================================
Sub Create_KSU(): Create_New_Row "КСУ АК", "Работа КСУ": End Sub
Sub Create_SU():  Create_New_Row "СУ АК", "Работа СУ":   End Sub
Sub Create_CNC(): Create_New_Row "Группа ЧПУ", "Работа ЧПУ": End Sub

Sub Create_New_Row(deptCode As String, workDesc As String)
    Dim n As clsOrder: Set n = PrepareOrder
    If n Is Nothing Then Exit Sub
    
    n.CreateNew deptCode, workDesc
    Show_Ch_Hint n.Department
End Sub

Sub Show_Ch_Hint(ByVal deptName As String)
    Dim wsRef As Worksheet: Set wsRef = ThisWorkbook.Sheets("Ref_Data")
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim f As Range
    
    wsP.Unprotect
    wsP.Range("PZ_DeptCode").ClearContents ' Очищаем старый код по ИМЕНИ
    
    Set f = wsRef.Columns("G").Find(What:=deptName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        Dim code As String: code = f.Offset(0, -1).Value
        wsP.Range("PZ_DeptCode").Value = code ' Пишем код по ИМЕНИ
        Application.StatusBar = "MES: Заказ создан! Код для ПЗ: " & code
    Else
        Application.StatusBar = "MES: Код для цеха '" & deptName & "' не найден"
    End If
    wsP.Protect
End Sub

' =========================================================
' БЛОК 3: БЕЗОПАСНАЯ ОТПРАВКА ПЗ (УМНЫЙ ЛОКАТОР)
' =========================================================
Sub PZ_SendToBase_Safe()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    
    ' 1. Считываем данные
    Dim sID As String: sID = UCase(wsP.Range("PZ_OrderNum").Text & wsP.Range("PZ_OrderPref").Text)
    Dim sDept As String: sDept = Trim(wsP.Range("PZ_Dept").Text)
    Dim sExtra As String: sExtra = UCase(Trim(wsP.Range("PZ_Extra").Text))
    Dim pzNum As String: pzNum = Trim(wsP.Range("PZ_Num").Text)
    
    If pzNum = "" Then MsgBox "Введите номер ПЗ!", 48: Exit Sub
    
    ' 2. ПОДКЛЮЧЕНИЕ К БАЗЕ
    Dim wbName As String: wbName = Trim(wsP.Range("PZ_DBName").Text)
    Dim wsB As Worksheet: On Error Resume Next
    Set wsB = Workbooks(wbName).Sheets(1)
    On Error GoTo 0
    
    If wsB Is Nothing Then
        MsgBox "База НзП (" & wbName & ") не найдена или закрыта!", 16: Exit Sub
    End If
    
    ' --- БРОНЕЖИЛЕТ ОТ READ-ONLY ---
    If wsB.Parent.ReadOnly Then
        MsgBox "База НзП открыта 'Только для чтения'! Отправка ПЗ заблокирована.", vbCritical, "MES: Ошибка доступа"
        Exit Sub
    End If
    ' -------------------------------
    
    ' 3. УМНЫЙ ПОИСК (Мягкое совпадение по частям)
    Dim i As Long, targetR As Long: targetR = 0
    Dim foundOrderButNotEmpty As Boolean: foundOrderButNotEmpty = False
    
    For i = wsB.Cells(wsB.Rows.count, 15).End(xlUp).Row To 2 Step -1
        ' Защита от ячеек с ошибками (#Н/Д, #ССЫЛКА!)
        If Not IsError(wsB.Cells(i, 15).Value) Then
            Dim cellName As String
            cellName = UCase(Trim(CStr(wsB.Cells(i, 15).Value2)))
            
            ' Ищем наличие номера заказа И цеха внутри ячейки
            If InStr(1, cellName, sID, vbTextCompare) > 0 And InStr(1, cellName, sDept, vbTextCompare) > 0 Then
                
                ' Если есть Приписка ОГЭ, она тоже должна быть внутри
                If sExtra = "" Or InStr(1, cellName, sExtra, vbTextCompare) > 0 Then
                
                    ' Проверяем, пуст ли ПЗ
                    If Trim(CStr(wsB.Cells(i, 2).Value2)) = "" Then
                        targetR = i
                        Exit For
                    Else
                        foundOrderButNotEmpty = True ' Нашли, но ПЗ уже занят
                    End If
                    
                End If
            End If
        End If
    Next i
    
    ' 4. ПОДТВЕРЖДЕНИЕ И ЗАПИСЬ
    If targetR > 0 Then
        ' Считываем реальные данные для контроля
        Dim realOrder As String: realOrder = wsB.Cells(targetR, 15).Value
        Dim realSection As String: realSection = wsB.Cells(targetR, 7).Value
        
        Dim promptMsg As String
        promptMsg = "ВНИМАНИЕ! Проверка адреса перед записью:" & vbCrLf & _
                    "------------------------------------------------" & vbCrLf & _
                    "Заказ: " & realOrder & vbCrLf & _
                    "Участок: " & realSection & vbCrLf & _
                    "Номер ПЗ: " & pzNum & vbCrLf & _
                    "------------------------------------------------" & vbCrLf & _
                    "Подтверждаете вброс в базу?"

        If MsgBox(promptMsg, vbQuestion + vbYesNo, "Контроль: ПЗ -> НзП") = vbNo Then
            Application.StatusBar = "MES: Вброс ПЗ отменен."
            Exit Sub
        End If

        Application.ScreenUpdating = False
        wsB.Cells(targetR, 2).Value = pzNum
        
        Run_Smart_Backup_Logic
        
        wsB.Parent.Save ' <--- ИСПРАВЛЕНО: Сохраняем саму базу НзП!
        ThisWorkbook.Save ' Сохраняем пульт, чтобы он запомнил очистку полей
        
        wsP.Unprotect
        wsP.Range("PZ_ItemCode, PZ_DeptCode, PZ_Num").ClearContents
        wsP.Protect
        Application.ScreenUpdating = True
        
        Application.StatusBar = "MES: ПЗ " & pzNum & " успешно записан"
        MsgBox "Готово! Данные отправлены в общую базу и синхронизированы.", 64
    Else
        ' Вывод точной причины ошибки
        If foundOrderButNotEmpty Then
            MsgBox "Заказ '" & sID & " " & sDept & "' найден, но в нём больше нет пустых ячеек для ПЗ!", 16
        Else
            MsgBox "Строка для заказа '" & sID & " " & sDept & "' вообще не найдена в базе!", 16
        End If
    End If
End Sub

Sub Undo_Last_Action()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim uRow As Long, uWB As String, uID As String, uDept As String
    
    ' Считываем следы
    uRow = val(wsP.Range("UNDO_Row").Value)
    uWB = wsP.Range("UNDO_WB").Text
    uID = wsP.Range("UNDO_ID").Text
    uDept = wsP.Range("UNDO_Dept").Text
    
    If uRow = 0 Or uWB = "" Then
        MsgBox "Нет данных для отмены!", vbExclamation, "MES: Отмена"
        Exit Sub
    End If
    
    Dim wsB As Worksheet
    On Error Resume Next
    Set wsB = Workbooks(uWB).Sheets(1)
    On Error GoTo 0
    
    If wsB Is Nothing Then Exit Sub
    
    ' Проверка: не изменилась ли строка (защита от удаления чужой работы)
    If Trim(wsB.Cells(uRow, 15).Text) = uID And Trim(wsB.Cells(uRow, 7).Text) = uDept And wsB.Cells(uRow, 2).Value = "" Then
        wsB.Rows(uRow).Delete
        MsgBox "Строка успешно удалена!", vbInformation, "MES: Отмена"
        
        ' Стираем память
        wsP.Range("UNDO_Row, UNDO_WB, UNDO_ID, UNDO_Dept").ClearContents
        wsB.Parent.Save
    Else
        MsgBox "Отмена невозможна! Строка была изменена или перемещена.", vbCritical, "MES: Защита"
    End If
End Sub

