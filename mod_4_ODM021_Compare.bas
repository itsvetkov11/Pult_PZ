Attribute VB_Name = "mod_4_ODM021_Compare"
Option Explicit

' Главный макрос для сравнения отчета ODM021 с листом НзП
Sub CompareODM021()
    Dim wsSettings As Worksheet
    Dim wsControl As Worksheet
    Dim wsNzP As Worksheet
    Dim wsReport As Worksheet
    Dim folderPath As String
    Dim latestFile As String
    Dim maxDate As Date
    Dim currentFile As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim wbName As String
    
    ' 1. Получение пути из настроек
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    
    If wsSettings Is Nothing Then
        MsgBox "Не найден лист 'Settings' в текущей книге.", vbCritical
        Exit Sub
    End If
    
    On Error Resume Next
    folderPath = wsSettings.Range("Путь_ODM021").Value
    On Error GoTo 0
    
    If folderPath = "" Then
        ' Попытка найти по тексту, если именованный диапазон не сработал
        Dim foundCell As Range
        Set foundCell = wsSettings.Columns("H").Find(What:="Путь_ODM021", LookIn:=xlValues, LookAt:=xlWhole)
        wsSettings.Cells.Find What:="", LookAt:=xlPart ' Сброс поиска (Ctrl+F) на частичное совпадение
        If Not foundCell Is Nothing Then
            folderPath = foundCell.Offset(0, 1).Value
        End If
    End If
    
    If folderPath = "" Then
        MsgBox "Не удалось найти путь к папке с отчетами ODM021 в настройках.", vbCritical
        Exit Sub
    End If
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' 2. Поиск последнего файла
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Папка " & folderPath & " не существует.", vbCritical
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(folderPath)
    maxDate = DateSerial(1900, 1, 1)
    
    For Each file In folder.Files
        If InStr(1, file.Name, "ФА_ODM021", vbTextCompare) > 0 And (Right(file.Name, 4) = ".xls" Or Right(file.Name, 5) = ".xlsx") Then
            If file.DateLastModified > maxDate Then
                maxDate = file.DateLastModified
                latestFile = file.Path
            End If
        End If
    Next file
    
    If latestFile = "" Then
        MsgBox "В папке " & folderPath & " не найдено отчетов ODM021.", vbExclamation
        Exit Sub
    End If
    
    ' 3. Подключение к базе НзП
    On Error Resume Next
    Set wsControl = ThisWorkbook.Sheets("PZ_Control")
    On Error GoTo 0
    
    If wsControl Is Nothing Then
        MsgBox "Не найден лист 'PZ_Control'.", vbCritical
        Exit Sub
    End If
    
    wbName = Trim(wsControl.Range("PZ_DBName").Text)
    
    On Error Resume Next
    Set wsNzP = Workbooks(wbName).Sheets(1)
    On Error GoTo 0
    
    If wsNzP Is Nothing Then
        MsgBox "База НзП (" & wbName & ") не найдена или закрыта! Пожалуйста, откройте базу перед запуском.", vbCritical
        Exit Sub
    End If
    
    If wsNzP.Parent.ReadOnly Then
        MsgBox "База НзП открыта 'Только для чтения'! Обновление дат заблокировано.", vbCritical, "MES: Ошибка доступа"
        Exit Sub
    End If
    
    ' 4. Подготовка словаря для быстрого поиска
    Dim dictNzP As Object
    Set dictNzP = CreateObject("Scripting.Dictionary")
    
    ' Поиск колонок в НзП (в 1-й строке)
    Dim colNzP_PZ As Long
    Dim colNzP_DateStatus As Long
    Dim colNzP_DateUpdate As Long
    Dim lastColNzP As Long
    Dim lastRowNzP As Long
    lastColNzP = wsNzP.Cells(1, wsNzP.Columns.Count).End(xlToLeft).Column
    
    Dim c As Long
    For c = 1 To lastColNzP
        If wsNzP.Cells(1, c).Value = "№ ПЗ" Then colNzP_PZ = c
        If wsNzP.Cells(1, c).Value = "Дата присвоения статуса" Then colNzP_DateStatus = c
        If wsNzP.Cells(1, c).Value = "Дата последнего обновления ПЗ" Then colNzP_DateUpdate = c
    Next c
    
    If colNzP_PZ = 0 Then
        MsgBox "На листе базы 'НзП' не найдена колонка '№ ПЗ' в 1-й строке.", vbCritical
        Exit Sub
    End If
    
    lastRowNzP = wsNzP.Cells(wsNzP.Rows.Count, colNzP_PZ).End(xlUp).Row
    
    ' Заполнение словаря
    Dim i As Long
    Dim pzVal As String
    For i = 2 To lastRowNzP
        pzVal = CStr(wsNzP.Cells(i, colNzP_PZ).Value)
        If pzVal <> "" And Not dictNzP.Exists(pzVal) Then
            dictNzP.Add pzVal, i
        End If
    Next i
    
    ' 5. Подготовка листа вывода
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Отчет_021")
    On Error GoTo 0
    
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReport.Name = "Отчет_021"
    Else
        wsReport.Cells.Clear
    End If
    
    ' 6. Открытие отчета и сравнение
    Application.ScreenUpdating = False
    Dim wbReport As Workbook
    Set wbReport = Workbooks.Open(latestFile)
    Dim wsRepData As Worksheet
    Set wsRepData = wbReport.Sheets(1) ' Предполагаем, что данные на первом листе
    
    ' Поиск колонок в отчете (в 8 строке)
    Dim colRep_PZ As Long
    Dim colRep_Dept As Long
    Dim colRep_DateStatus As Long
    Dim colRep_DateUpdate As Long
    Dim lastColRep As Long
    lastColRep = wsRepData.Cells(8, wsRepData.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastColRep
        If wsRepData.Cells(8, c).Value = "№ ПЗ" Then colRep_PZ = c
        If wsRepData.Cells(8, c).Value = "Отдел" Then colRep_Dept = c
        If wsRepData.Cells(8, c).Value = "Дата присвоения статуса" Then colRep_DateStatus = c
        If wsRepData.Cells(8, c).Value = "Дата последнего обновления ПЗ" Then colRep_DateUpdate = c
    Next c
    
    If colRep_PZ = 0 Then
        MsgBox "В отчете не найдена колонка '№ ПЗ' в 8-й строке.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    If colRep_Dept = 0 Then
        MsgBox "В отчете не найдена колонка 'Отдел' в 8-й строке.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Копирование шапки
    wsRepData.Rows(8).Copy Destination:=wsReport.Rows(1)
    
    Dim lastRowRep As Long
    lastRowRep = wsRepData.Cells(wsRepData.Rows.Count, colRep_PZ).End(xlUp).Row
    
    Dim outRow As Long
    outRow = 2
    Dim deptVal As String
    Dim nzpRow As Long
    Dim updatedCount As Long
    updatedCount = 0
    
    wsNzP.Unprotect Password:="1"
    For i = 9 To lastRowRep
        deptVal = Trim(CStr(wsRepData.Cells(i, colRep_Dept).Value))
        ' Фильтрация по отделу
        If deptVal = "СУ АК" Or deptVal = "КСУ АК" Or deptVal = "Группа ЧПУ" Then
            pzVal = CStr(wsRepData.Cells(i, colRep_PZ).Value)
            If pzVal <> "" Then
                If Not dictNzP.Exists(pzVal) Then
                    ' Строки нет в НзП, копируем в отчет
                    wsRepData.Rows(i).Copy Destination:=wsReport.Rows(outRow)
                    outRow = outRow + 1
                Else
                    ' Строка найдена в НзП - обновляем даты, если колонки найдены
                    nzpRow = dictNzP(pzVal)
                    If colRep_DateStatus > 0 And colNzP_DateStatus > 0 Then
                        wsNzP.Cells(nzpRow, colNzP_DateStatus).Value = wsRepData.Cells(i, colRep_DateStatus).Value
                    End If
                    If colRep_DateUpdate > 0 And colNzP_DateUpdate > 0 Then
                        wsNzP.Cells(nzpRow, colNzP_DateUpdate).Value = wsRepData.Cells(i, colRep_DateUpdate).Value
                    End If
                    updatedCount = updatedCount + 1
                End If
            End If
        End If
    Next i
    
    wsNzP.Protect Password:="1", AllowFiltering:=True
    wbReport.Close SaveChanges:=False
    
    If updatedCount > 0 Then
        On Error Resume Next
        wsNzP.Parent.Save
        On Error GoTo 0
    End If
    
    wsReport.Activate
    Application.ScreenUpdating = True
    
    MsgBox "Готово! Обработан файл: " & vbCrLf & latestFile & vbCrLf & _
           "Найдено отсутствующих строк: " & (outRow - 2) & vbCrLf & _
           "Обновлено дат в базе НзП: " & updatedCount, vbInformation
    
End Sub

