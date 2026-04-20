VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject: On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Set tbl = ws.ListObjects("tblZVR_Master")
        If Not tbl Is Nothing Then Exit For
    Next ws: On Error GoTo 0
    
    If tbl Is Nothing Then Exit Sub
    On Error GoTo CleanExit

    ' --- А. ВЫБОР ИЗ ИСТОРИИ (Телепорт - B2 оставляем как есть) ---
    If Not Intersect(Target, Me.Range("B2")) Is Nothing Then
        If Me.Range("B2").Value <> "" Then
            Application.EnableEvents = False
            UpdateSearchHistory Me.Range("B2").Value
            ' Передаем заказ в поле поиска по ИМЕНИ
            Me.Range("PZ_SearchOrder").Value = Me.Range("B2").Value
            Application.EnableEvents = True
            SearchByOrder Me.Range("PZ_SearchOrder"), tbl
        End If
    End If

    ' --- Б. ПОИСК ПО ЗВР (Работает по ИМЕНИ) ---
    If Not Intersect(Target, Me.Range("PZ_SearchZVR")) Is Nothing Then
        SearchByZVR Target, tbl
    End If

    ' --- В. ПОИСК ПО ЗАКАЗУ (Работает по ИМЕНИ) ---
    If Not Intersect(Target, Me.Range("PZ_SearchOrder")) Is Nothing Then
        SearchByOrder Target, tbl
    End If
    
    ' --- Г. АВТО-ПРЕФИКС "ФК" (Работает по ИМЕНИ PZ_OrderNum) ---
    If Not Intersect(Target, Me.Range("PZ_OrderNum")) Is Nothing Then
        Application.EnableEvents = False
        If Right(Target.Text, 1) = "-" Then Me.Range("PZ_OrderPref").Value = "ФК"
        Application.EnableEvents = True
    End If

CleanExit:
    Application.EnableEvents = True
End Sub

' --- ВСПОМОГАТЕЛЬНЫЕ ПОДПРОГРАММЫ (Поиск) ---

Sub SearchByZVR(ByVal Target As Range, tbl As ListObject)
    Dim sSearch As String: sSearch = Trim(Target.Text)
    Dim foundRow As Range: If sSearch = "" Or sSearch = "Не найдена" Then Exit Sub
    
    Application.EnableEvents = False
    Me.Range("PZ_SearchOrder, PZ_SearchClient").ClearContents ' Чистим по ИМЕНАМ
    
    Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(What:=sSearch, LookIn:=xlValues, LookAt:=xlWhole)
    If foundRow Is Nothing And IsNumeric(sSearch) Then
        Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(What:=CLng(sSearch), LookIn:=xlValues, LookAt:=xlWhole)
    End If
    
    If Not foundRow Is Nothing Then
        Me.Range("PZ_SearchOrder").Value = foundRow.Offset(0, 1).Value
        Me.Range("PZ_SearchClient").Value = foundRow.Offset(0, 2).Value
        UpdateSearchHistory Me.Range("PZ_SearchOrder").Value
    Else
        Me.Range("PZ_SearchOrder").Value = "Не найден"
    End If
    
    tbl.DataBodyRange.Find What:="", LookAt:=xlPart
    Application.EnableEvents = True
End Sub

Sub SearchByOrder(ByVal Target As Range, tbl As ListObject)
    Dim sSearch As String: sSearch = Trim(Target.Text)
    Dim foundRow As Range: If sSearch = "" Or sSearch = "Не найден" Then Exit Sub
    
    Application.EnableEvents = False
    Me.Range("PZ_SearchZVR, PZ_SearchClient").ClearContents ' Чистим по ИМЕНАМ
    UpdateSearchHistory sSearch
    
    Set foundRow = tbl.ListColumns(2).DataBodyRange.Find(What:=sSearch, LookIn:=xlValues, LookAt:=xlWhole)
    If foundRow Is Nothing And IsNumeric(sSearch) Then
        Set foundRow = tbl.ListColumns(2).DataBodyRange.Find(What:=CLng(sSearch), LookIn:=xlValues, LookAt:=xlWhole)
    End If
    
    If Not foundRow Is Nothing Then
        Me.Range("PZ_SearchZVR").Value = foundRow.Offset(0, -1).Value
        Me.Range("PZ_SearchClient").Value = foundRow.Offset(0, 1).Value
    Else
        Me.Range("PZ_SearchZVR").Value = "Не найдена"
    End If
    
    tbl.DataBodyRange.Find What:="", LookAt:=xlPart
    Application.EnableEvents = True
End Sub
