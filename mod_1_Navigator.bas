Attribute VB_Name = "mod_1_Navigator"
Option Explicit

' =========================================================
' БЛОК 1: НАВИГАЦИЯ И ЗАКРЫТИЕ (Работает через PZ_SearchMain)
' =========================================================

Function CurrentOrder() As clsOrder
    Dim ord As New clsOrder
    ord.InitializeSearch ThisWorkbook.Sheets("PZ_Control").Range("PZ_SearchMain").Text
    Set CurrentOrder = ord
End Function

Sub PZ_Teleport()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim fVal As String: fVal = Trim(wsP.Range("PZ_SearchMain").Text)
    
    If fVal = "" Then MsgBox "Введите номер в поле поиска!", 48: Exit Sub
    
    UpdateSearchHistory fVal
    
    Dim ord As clsOrder: Set ord = CurrentOrder
    If ord.Rows.count = 0 Then
        Application.StatusBar = "MES: Заказ " & fVal & " не найден"
        MsgBox "Заказ '" & fVal & "' не найден!", 48: Exit Sub
    End If
    
    Dim idx As Long: idx = val(wsP.Range("PZ_TeleportIdx").Value) + 1
    If idx > ord.Rows.count Then idx = 1
    
    wsP.Unprotect
    wsP.Range("PZ_TeleportIdx").Value = idx
    wsP.Protect UserInterfaceOnly:=True
    
    Dim targetRow As Long: targetRow = ord.Rows(idx)
    Dim wsB As Worksheet: Set wsB = ord.BaseSheet
    
    Application.ScreenUpdating = True
    wsB.Parent.Activate: wsB.Activate
    On Error Resume Next: AppActivate wsB.Parent.Name: On Error GoTo 0
    DoEvents
    
    wsB.Cells(targetRow, 15).Select
    With ActiveWindow
        .ScrollRow = targetRow
        .SmallScroll Down:=1: .SmallScroll Up:=1
    End With
    
    Application.StatusBar = "MES Телепорт " & fVal & ": " & idx & " из " & ord.Rows.count
    
    ' ДОПОЛНИТЕЛЬНЫЙ ПОИСК ПО ЗВР
    Dim zvrVal As String, zvrSearchTerm As String
    On Error Resume Next
    zvrVal = Trim(wsP.Range("PZ_SearchZVR").Text)
    On Error GoTo 0
    
    If zvrVal <> "" And zvrVal <> "Не найден" And zvrVal <> "Не найдена" And zvrVal <> fVal Then
        ' Если ЗВР не начинается с дефиса, добавляем его, 
        ' чтобы IsMatch сделал поиск вхождения без строгих границ
        If Left(zvrVal, 1) <> "-" Then
            zvrSearchTerm = "-" & zvrVal
        Else
            zvrSearchTerm = zvrVal
        End If
        
        Dim zvrOrd As New clsOrder
        zvrOrd.InitializeSearch zvrSearchTerm
        
        If zvrOrd.Rows.count > 0 Then
            Dim rowList As String
            Dim i As Long
            rowList = ""
            For i = 1 To zvrOrd.Rows.count
                rowList = rowList & zvrOrd.Rows(i)
                If i < zvrOrd.Rows.count Then rowList = rowList & ", "
            Next i
            MsgBox "По номеру ЗВР (" & zvrVal & ") найдены дополнительные строки в основной таблице." & vbCrLf & "Номера строк: " & rowList, vbInformation, "Дополнительный поиск по ЗВР"
        End If
    End If
End Sub

Sub PZ_ProcessRow()
    CurrentOrder.ApplyStyling
    MsgBox "Готово!", 64
End Sub

' МОСТЫ: Добавить участок в СУЩЕСТВУЮЩИЙ заказ (Закрытие)
Sub Add_KSU(): CurrentOrder.AddSection "КСУ АК", "Работа КСУ": End Sub
Sub Add_SU():  CurrentOrder.AddSection "СУ АК", "Работа СУ":   End Sub
Sub Add_CNC(): CurrentOrder.AddSection "Группа ЧПУ", "Работа ЧПУ": End Sub

' =========================================================
' БЛОК 2: ХИРУРГИЧЕСКАЯ ИСТОРИЯ (Без сдвига ячеек)
' =========================================================
Sub UpdateSearchHistory(ByVal newVal As String)
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim histRange As Range: Set histRange = wsP.Range("PZ_SearchHistory")
    Dim i As Integer
    
    If newVal = "" Or newVal = "Не найден" Or newVal = "Не найдена" Then Exit Sub
    
    Application.EnableEvents = False
    wsP.Unprotect
    
    Dim mIdx As Variant
    mIdx = Application.Match(newVal, histRange, 0)
    
    If Not IsError(mIdx) Then
        For i = mIdx To 2 Step -1 ' Изменено для работы с Range напрямую, если он отвязан от колонок. Но оставим логику сдвига через Cells, если она жестко привязана.
            histRange.Cells(i, 1).Value = histRange.Cells(i - 1, 1).Value
        Next i
    Else
        For i = 10 To 2 Step -1
            histRange.Cells(i, 1).Value = histRange.Cells(i - 1, 1).Value
        Next i
    End If
    
    histRange.Cells(1, 1).Value = newVal
    wsP.Protect UserInterfaceOnly:=True
    Application.EnableEvents = True
End Sub
