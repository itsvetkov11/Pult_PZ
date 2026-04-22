VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrder"
Option Explicit

Public Rows As Collection
Public SearchTerm As String
Public ID As String
Public Department As String
Public WorkType As String
Public ExtraInfo As String
Public ItemCode As String
Public Description As String

Property Get BaseSheet() As Worksheet
    Dim wbName As String
    wbName = Trim(ThisWorkbook.Sheets("PZ_Control").Range("PZ_DBName").Text)
    On Error Resume Next
    Set BaseSheet = Workbooks(wbName).Sheets(1)
    On Error GoTo 0
End Property

Private Sub Class_Initialize()
    Set Rows = New Collection
End Sub

' =========================================================
' МАГИЯ ПРОРИСОВКИ И ФОКУСА НА СТРОКЕ
' =========================================================
Private Sub ForceFocus(ws As Worksheet, targetRow As Long)
    On Error Resume Next
    ws.Parent.Activate
    ws.Activate
    ws.Cells(targetRow, 15).Select
    With ActiveWindow
        .ScrollRow = targetRow ' Прижимаем строку к шапке
        .SmallScroll Down:=1: .SmallScroll Up:=1 ' Микро-скролл для графики
    End With
    ' Эмуляция переключения окон (Alt+Tab эффект)
    AppActivate ws.Parent.Name
    DoEvents
End Sub

' =========================================================
' ПОИСК (ОРИГИНАЛ v6.0)
' =========================================================
Sub InitializeSearch(ByVal term As String)
    Dim wsB As Worksheet: Set wsB = Me.BaseSheet
    If wsB Is Nothing Then Exit Sub

    Me.SearchTerm = term
    Set Rows = New Collection

    Dim i As Long, lastR As Long, txt As String
    lastR = wsB.Cells(wsB.Rows.count, 15).End(xlUp).Row

    For i = 2 To lastR
        If wsB.Cells(i, 15).MergeCells Then
            txt = wsB.Cells(i, 15).MergeArea.Cells(1, 1).Text
        Else
            txt = wsB.Cells(i, 15).Text
        End If

        If IsMatch(txt) Then
            If Trim(wsB.Cells(i, 7).Value) <> "" Then Rows.Add i
        End If
    Next i
End Sub

Private Function IsMatch(txt As String) As Boolean
    Dim isRV As Boolean, actual As String, pos As Long
    isRV = (Left(SearchTerm, 1) = "-")
    actual = IIf(isRV, Mid(SearchTerm, 2), SearchTerm)
    pos = InStr(1, txt, actual, vbTextCompare)

    If pos = 0 Then Exit Function
    If isRV Then IsMatch = True: Exit Function

    IsMatch = True
    If pos > 1 Then
        If IsNumeric(Mid(txt, pos - 1, 1)) Or Mid(txt, pos - 1, 1) = "-" Then IsMatch = False
    End If
    If IsMatch And (pos + Len(actual)) <= Len(txt) Then
        If IsNumeric(Mid(txt, pos + Len(actual), 1)) Then IsMatch = False
    End If
End Function

Sub LoadFromRegistry()
    Dim wsRef As Worksheet: Set wsRef = ThisWorkbook.Sheets("Ref_Data")
    Dim tbl As ListObject: Set tbl = wsRef.ListObjects("tblItems")
    Dim r As Range, pref As String, searchDept As String

    searchDept = Me.Department
    If searchDept = "ЦПМ-4" Then searchDept = "ЦПМ-2"

    Dim isServiceDept As Boolean
    isServiceDept = Not IsError(Application.Match(searchDept, Array("0", "ВФ", "КФ", "БФ", "Сторонние"), 0))

    If LCase(Me.WorkType) = "услуги" Or isServiceDept Then
        Me.ItemCode = "32-000000-00004"
        Me.Description = "Услуги производственно-технического назначения"
        Exit Sub
    End If

    pref = IIf(LCase(Me.WorkType) = "до", "32-291", "32-282")

    For Each r In tbl.ListColumns(1).DataBodyRange
        If Left(r.Value, 6) = pref Then
            If InStr(1, r.Offset(0, 1).Value, searchDept, vbTextCompare) > 0 Then
                Me.ItemCode = r.Value
                Me.Description = r.Offset(0, 1).Value
                Exit Sub
            End If
        End If
    Next r
    Me.ItemCode = ""
End Sub

' =========================================================
' СОЗДАНИЕ СТРОКИ (+ ДОБАВЛЕН ФОКУС И СОХРАНЕНИЕ)
' =========================================================
Sub CreateNew(ByVal deptName As String, ByVal workDesc As String)
    Dim wsB As Worksheet: Set wsB = Me.BaseSheet
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim fullName As String, targetR As Long, i As Long

    If wsB Is Nothing Then Exit Sub

    ' --- БРОНЕЖИЛЕТ ОТ READ-ONLY ---
    If wsB.Parent.ReadOnly Then
        MsgBox "База НзП открыта 'Только для чтения'! (Нормально зайди в таблицу, нормально будет!)))." & vbCrLf & "Внесение изменений заблокировано.", vbCritical, "MES: Ошибка доступа"
        Exit Sub
    End If
    ' -------------------------------

    Me.ID = UCase(Me.ID)
    Me.ExtraInfo = UCase(Me.ExtraInfo)

    Dim searchPrefix As String
    Select Case LCase(Me.WorkType)
        Case "до":      searchPrefix = "32-291"
        Case "но":      searchPrefix = "32-282"
        Case "услуги":  searchPrefix = "32-000"
        Case Else:      searchPrefix = "32-282"
    End Select

    targetR = 0
    Dim lastR As Long: lastR = wsB.Cells(wsB.Rows.count, 15).End(xlUp).Row
    Dim foundFullName As String: foundFullName = ""

    ' ОПТИМИЗАЦИЯ: Загружаем диапазон в массив
    Dim vData As Variant
    vData = wsB.Range("A1:O" & lastR).Value2

    For i = lastR To 2 Step -1
        If Not IsError(vData(i, 15)) And Not IsError(vData(i, 5)) Then
            Dim cellTxt As String: cellTxt = UCase(Trim(CStr(vData(i, 15))))
            If InStr(1, cellTxt, Me.ID, vbTextCompare) > 0 And InStr(1, cellTxt, Me.Department, vbTextCompare) > 0 Then
                If Left(CStr(vData(i, 5)), 6) = searchPrefix Then
                    targetR = i + 1
                    foundFullName = Trim(CStr(vData(i, 15)))
                    Exit For
                End If
            End If
        End If
    Next i

    If foundFullName <> "" Then
        fullName = foundFullName
    Else
        fullName = Me.ID
        If Me.ExtraInfo <> "" Then fullName = fullName & " " & Me.ExtraInfo
        fullName = fullName & " " & Me.Department
    End If

    ' Проверка на дубликат + ТЕЛЕПОРТ К ОШИБКЕ
        For i = 2 To lastR
        If Not IsError(vData(i, 15)) And Not IsError(vData(i, 7)) Then
            If Trim(CStr(vData(i, 15))) = fullName And Trim(CStr(vData(i, 7))) = deptName Then
                MsgBox "Ошибка! Участок '" & deptName & "' уже добавлен в заказ " & fullName & "!", 48, "Дубликат"
                ForceFocus wsB, i ' Фокусируемся на дубликате
                Exit Sub
            End If
        End If
    Next i

    If targetR = 0 Then
        For i = lastR To 2 Step -1
            If Not IsError(vData(i, 15)) And Not IsError(vData(i, 5)) Then
                If InStr(1, CStr(vData(i, 15)), Me.Department, vbTextCompare) > 0 And Left(CStr(vData(i, 5)), 6) = searchPrefix Then
                    targetR = i + 1: Exit For
                End If
            End If
        Next i
    End If

    If targetR = 0 Then
        For i = lastR To 2 Step -1
            If Not IsError(vData(i, 15)) Then
                If InStr(1, CStr(vData(i, 15)), Me.Department, vbTextCompare) > 0 Then
                    targetR = i + 1: Exit For
                End If
            End If
        Next i
    End If

    If targetR = 0 Then targetR = lastR + 1

    Application.ScreenUpdating = False
    wsB.Rows(targetR).Insert Shift:=xlDown

    With wsB.Rows(targetR)
        .Interior.ColorIndex = xlNone: .Font.ColorIndex = xlAutomatic
    End With

    With wsB
        .Cells(targetR, 2).Value = ""
        .Cells(targetR, 3).Value = "МЧ1"
        .Cells(targetR, 4).Value = "Не выпущено"
        .Cells(targetR, 5).Value = Me.ItemCode
        .Cells(targetR, 6).Value = Me.Description
        .Cells(targetR, 7).Value = deptName
        .Cells(targetR, 8).Value = Date
        .Cells(targetR, 9).Value = ""
        .Cells(targetR, 15).Value = fullName
        .Cells(targetR, 16).Value = workDesc

        With .Cells(targetR, 17)
            .Value = "списание": .Font.Color = vbRed
        End With

        .Cells(targetR, 21).Value = ""

        With .Range(.Cells(targetR, 1), .Cells(targetR, 22)).Borders
            .LineStyle = xlContinuous: .Weight = xlThin
        End With
    End With

    wsP.Unprotect
    wsP.Range("PZ_ItemCode").Value = Me.ItemCode

    ' ЗАПИСЬ ДЛЯ УМНОГО UNDO
    wsP.Range("UNDO_Row").Value = targetR
    wsP.Range("UNDO_WB").Value = wsB.Parent.Name
    wsP.Range("UNDO_ID").Value = fullName
    wsP.Range("UNDO_Dept").Value = deptName
    wsP.Protect

    ' СИНХРОНИЗАЦИЯ С ОБЩЕЙ БАЗОЙ
    wsB.Parent.Save

    Application.ScreenUpdating = True
    ForceFocus wsB, targetR ' Фокус на созданной строке
End Sub

' =========================================================
' ПОКРАСКА (+ ДОБАВЛЕНЫ ПОЛНЫЕ РАМКИ СО ВСЕХ СТОРОН)
' =========================================================
Sub ApplyStyling()
    If Rows.count = 0 Then Exit Sub
    Dim wsB As Worksheet: Set wsB = Me.BaseSheet
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim f2 As Range: Set f2 = wsP.Range("PZ_StyleColor")

    If wsB Is Nothing Then Exit Sub ' Добавим на всякий случай

    ' --- БРОНЕЖИЛЕТ ОТ READ-ONLY ---
    If wsB.Parent.ReadOnly Then
        MsgBox "База НзП открыта 'Только для чтения'! Покраска заблокирована.", vbCritical, "MES: Ошибка доступа"
        Exit Sub
    End If
    ' -------------------------------

    Me.GatherRows

    Dim colUnique As New Collection
    Dim rNum As Variant, fullName As String
    On Error Resume Next
    For Each rNum In Rows
        fullName = Trim(wsB.Cells(val(rNum), 15).Text)
        colUnique.Add fullName, CStr(fullName)
    Next rNum
    On Error GoTo 0

    For Each rNum In Rows
        wsB.Cells(val(rNum), 21).Value = wsP.Range("PZ_StyleMark").Value
        wsB.Cells(val(rNum), 17).Value = f2.Value

        With wsB.Range(wsB.Cells(val(rNum), 1), wsB.Cells(val(rNum), 22))
            .Interior.Color = f2.Interior.Color
            .Font.Color = f2.Font.Color
        End With
    Next rNum

    Dim i As Integer, minR As Long, maxR As Long
    Dim edge As Variant

    For i = 1 To colUnique.count
        minR = 999999: maxR = 0

        For Each rNum In Rows
            If Trim(wsB.Cells(val(rNum), 15).Text) = colUnique(i) Then
                If val(rNum) < minR Then minR = val(rNum)
                If val(rNum) > maxR Then maxR = val(rNum)
            End If
        Next rNum

        If minR <= maxR Then
            With wsB.Range(wsB.Cells(minR, 15), wsB.Cells(maxR, 15))
                .Borders.LineStyle = xlNone

                ' Рисуем рамку по всему периметру
                For Each edge In Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
                    With .Borders(edge)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                        .Color = f2.Borders(xlEdgeLeft).Color
                    End With
                Next edge

                ' Тонкие линии внутри блока (если строк больше одной)
                If minR < maxR Then
                    With .Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End If
            End With
        End If
    Next i

    ' СИНХРОНИЗАЦИЯ С ОБЩЕЙ БАЗОЙ
    wsB.Parent.Save
    ForceFocus wsB, minR ' Фокус на покрашенном блоке
End Sub

' =========================================================
' ДОБАВЛЕНИЕ УЧАСТКА (+ ДОБАВЛЕН ФОКУС И СОХРАНЕНИЕ)
' =========================================================
Sub AddSection(deptName As String, workDesc As String)
    If Rows.count = 0 Then Exit Sub
    Dim wsB As Worksheet: Set wsB = Me.BaseSheet
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")

    If wsB Is Nothing Then Exit Sub

    ' --- БРОНЕЖИЛЕТ ОТ READ-ONLY ---
    If wsB.Parent.ReadOnly Then
        MsgBox "База НзП открыта 'Только для чтения'! Добавление участка заблокировано.", vbCritical, "MES: Ошибка доступа"
        Exit Sub
    End If
    ' -------------------------------

    Dim colUnique As New Collection
    Dim rNum As Variant, fullName As String
    On Error Resume Next
    For Each rNum In Rows
        fullName = Trim(wsB.Cells(val(rNum), 15).Text)
        colUnique.Add fullName, CStr(fullName)
    Next rNum
    On Error GoTo 0

    Dim targetFullName As String

    If colUnique.count = 1 Then
        targetFullName = colUnique(1)
    Else
        Dim promptStr As String
        promptStr = "Найдено несколько цехов для этого заказа. Куда добавляем участок?" & vbCrLf & vbCrLf
        Dim idx As Integer
        For idx = 1 To colUnique.count
            promptStr = promptStr & idx & " - " & colUnique(idx) & vbCrLf
        Next idx
        promptStr = promptStr & vbCrLf & "Введите НОМЕР нужного варианта (1-" & colUnique.count & "):"

        Dim choice As String
        choice = InputBox(promptStr, "MES: Уточнение адреса")

        If choice = "" Or Not IsNumeric(choice) Then
            Application.StatusBar = "MES: Добавление отменено."
            Exit Sub
        End If
        If val(choice) < 1 Or val(choice) > colUnique.count Then
            MsgBox "Неверный номер!", 16: Exit Sub
        End If
        targetFullName = colUnique(val(choice))
    End If

    Dim i As Long, existingR As Long, lastSubOrderRow As Long
    Dim templateR As Long: templateR = 0

    For i = 1 To Rows.count
        If Trim(wsB.Cells(Rows(i), 15).Text) = targetFullName Then
            If templateR = 0 Then templateR = Rows(i)
            If Rows(i) > lastSubOrderRow Then lastSubOrderRow = Rows(i)

            If Trim(wsB.Cells(Rows(i), 7).Text) = deptName Then
                existingR = Rows(i): Exit For
            End If
        End If
    Next i

    If existingR > 0 Then
        MsgBox "Участок '" & deptName & "' уже есть в заказе " & targetFullName & "!", 64
        ForceFocus wsB, existingR ' ТЕЛЕПОРТ К ОШИБКЕ
        Exit Sub
    End If

    Dim inHours As Variant: inHours = Application.InputBox("Введите часы:", "Ввод трудозатрат", Type:=2)
    If inHours = "False" Or inHours = "" Then Exit Sub

    Application.ScreenUpdating = False

    wsB.Rows(lastSubOrderRow + 1).Insert Shift:=xlDown
    wsB.Range(wsB.Cells(templateR, 1), wsB.Cells(templateR, 22)).Copy Destination:=wsB.Range("A" & lastSubOrderRow + 1)

    With wsB.Rows(lastSubOrderRow + 1)
        .Cells(1, 1).ClearContents
        .Cells(1, 2).ClearContents
        .Cells(1, 8).Resize(1, 2).ClearContents
        .Cells(1, 11).Resize(1, 4).ClearContents
        .Cells(1, 18).Resize(1, 3).ClearContents

        .Cells(1, 7).Value = deptName
        .Cells(1, 16).Value = workDesc
        .Cells(1, 10).Value = inHours
        .Cells(1, 10).Font.Color = vbRed
        .Cells(1, 21).ClearContents
    End With

    Me.InitializeSearch Me.SearchTerm
    Me.ApplyStyling

    ' СИНХРОНИЗАЦИЯ С ОБЩЕЙ БАЗОЙ
    wsB.Parent.Save

    Application.ScreenUpdating = True
    Application.StatusBar = "MES: Участок " & deptName & " добавлен в " & targetFullName
    ForceFocus wsB, lastSubOrderRow + 1 ' Фокус на добавленной строке

    ' ЗАПИСЬ ДЛЯ УМНОГО UNDO
    wsP.Unprotect
    wsP.Range("UNDO_Row").Value = lastSubOrderRow + 1
    wsP.Range("UNDO_WB").Value = wsB.Parent.Name
    wsP.Range("UNDO_ID").Value = targetFullName
    wsP.Range("UNDO_Dept").Value = deptName
    wsP.Protect
End Sub

' =========================================================
' СБОРКА СТРОК (ОРИГИНАЛ v6.0)
' =========================================================
Sub GatherRows()
    Dim wsB As Worksheet: Set wsB = Me.BaseSheet
    If wsB Is Nothing Then Exit Sub
    If Rows.count <= 1 Then Exit Sub

    Dim lastR As Long
    lastR = wsB.Cells(wsB.Rows.count, 15).End(xlUp).Row

    Dim r As Long, anchorR As Long, insertR As Long
    anchorR = 0


    For r = 2 To lastR
        If IsMatch(Trim(wsB.Cells(r, 15).Text)) Then
            anchorR = r
            insertR = anchorR + 1
            Exit For
        End If
    Next r

    If anchorR = 0 Then Exit Sub

    r = insertR
    Do While r <= lastR
        If IsMatch(Trim(wsB.Cells(r, 15).Text)) Then
            If r > insertR Then
                Application.ScreenUpdating = False
                wsB.Rows(r).Cut
                wsB.Rows(insertR).Insert Shift:=xlDown
                insertR = insertR + 1
            Else
                insertR = insertR + 1
                r = r + 1
            End If
        Else
            r = r + 1
        End If
    Loop

    Me.InitializeSearch Me.SearchTerm
End Sub
