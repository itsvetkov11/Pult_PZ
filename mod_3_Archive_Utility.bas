Attribute VB_Name = "mod_3_Archive_Utility"
Option Explicit

Sub Clear_Pulse()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("PZ_Control")
    
    Application.EnableEvents = False
    
    ws.Unprotect
    ' Чистим ВСЕ блоки исключительно по именам!
    ws.Range("PZ_OrderNum, PZ_OrderPref, PZ_Dept, PZ_WorkType, PZ_Extra").ClearContents
    ws.Range("PZ_ItemCode, PZ_DeptCode, PZ_Num").ClearContents
    ws.Range("PZ_SearchZVR, PZ_SearchOrder, PZ_SearchClient").ClearContents
    ws.Protect UserInterfaceOnly:=True
    
    Application.StatusBar = False
    Application.EnableEvents = True
    
    ' Возвращаем курсор в поле поиска ЗВР
    ws.Range("PZ_SearchZVR").Select
End Sub

Sub Update_Bases_Manual()
    Dim msg As String
    
    msg = "Согласно принципу разумной достаточности, человека мы пока заменять не будем! :)" & vbCrLf & vbCrLf & _
          "Пожалуйста, обновляйте базы штатным способом (как показано на картинке чуть ниже уведомления):" & vbCrLf & _
          "Вкладка 'Данные' -> 'Обновить всё'." & vbCrLf & vbCrLf & _
          "Это самый надежный вариант для работы в общей сети."
          
    MsgBox msg, vbInformation, "РМЦ: Инструкция по обновлению"
End Sub

' ГЛАВНЫЙ ДИСПЕТЧЕР БЭКАПОВ
Sub Run_Smart_Backup_Logic()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    
    ' 1. Бронежилет: Если Пульт открыт только для чтения, мы не сможем обновить даты
    If ThisWorkbook.ReadOnly Then Exit Sub
    
    Dim wbName As String: wbName = Trim(wsP.Range("PZ_DBName").Text)
    
    ' Проверяем, открыта ли база
    Dim wbBase As Workbook
    On Error Resume Next
    Set wbBase = Workbooks(wbName)
    On Error GoTo 0
    
    If wbBase Is Nothing Then Exit Sub ' Если база не открыта - выходим (решает баг холостого выстрела)
    If wbBase.ReadOnly Then Exit Sub   ' Если база открыта для чтения, мы не сможем обновить метку
    
    ' Получаем или создаем скрытый лист с метками внутри САМОЙ БАЗЫ
    Dim wsSys As Worksheet
    On Error Resume Next
    Set wsSys = wbBase.Sheets("Sys_Backup")
    On Error GoTo 0
    
    If wsSys Is Nothing Then
        Set wsSys = wbBase.Sheets.Add(After:=wbBase.Sheets(wbBase.Sheets.count))
        wsSys.Name = "Sys_Backup"
        wsSys.Visible = xlSheetVeryHidden ' Скрываем от глаз
        wsSys.Range("A1").Value = "Last_AM_Backup"
        wsSys.Range("A2").Value = "Last_11_Backup"
    End If
    
    Dim lastAM As Date: lastAM = val(wsSys.Range("B1").Value)
    Dim last11 As Date: last11 = val(wsSys.Range("B2").Value)
    Dim currDate As Date: currDate = Date
    Dim currHour As Integer: currHour = Hour(Now)
    Dim backupSuccess As Boolean: backupSuccess = False

    ' 2. Логика "Кто первый встал" (Утренний бэкап)
    If lastAM < currDate Then
        backupSuccess = Execute_Silent_Backup(wbBase, "AM")
        If backupSuccess Then
            wsSys.Range("B1").Value = currDate
            wbBase.Save
        End If
    
    ' 3. Логика "11-часовой чекпоинт"
    ElseIf currHour >= 11 And last11 < currDate Then
        backupSuccess = Execute_Silent_Backup(wbBase, "11AM")
        If backupSuccess Then
            wsSys.Range("B2").Value = currDate
            wbBase.Save
        End If
    End If

    If backupSuccess Then Clean_Old_Backups ' Запускаем санитара (7 дней) только если был бэкап
End Sub

' ТИХАЯ КОПИЯ ФАЙЛА
Private Function Execute_Silent_Backup(wbBase As Workbook, bType As String) As Boolean
    On Error Resume Next
    Execute_Silent_Backup = False
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sourcePath As String: sourcePath = wbBase.fullName
    Dim backupFolder As String: backupFolder = ThisWorkbook.Path & "\_MES_Backups\"
    
    ' Создаем папку, если её нет
    If Not fso.FolderExists(backupFolder) Then fso.CreateFolder (backupFolder)
    
    ' Формируем имя: База_тип_дата_время.xlsx
    Dim baseName As String: baseName = fso.GetBaseName(sourcePath)
    Dim destPath As String: destPath = backupFolder & baseName & "_" & bType & "_" & Format(Now, "dd-mm-yyyy_HH-mm") & ".xlsx"
    
    ' Копируем файл "на лету"
    fso.CopyFile sourcePath, destPath, True
    
    If Err.Number = 0 Then
        Execute_Silent_Backup = True
        Application.StatusBar = "MES: Создан резервный слепок базы (" & bType & ")"
    End If
End Function

' САНИТАР (Удаление старых файлов > 7 дней)
Private Sub Clean_Old_Backups()
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(ThisWorkbook.Path & "\_MES_Backups\")
    Dim file As Object
    
    For Each file In folder.Files
        If DateDiff("d", file.DateCreated, Now) > 7 And LCase(fso.GetExtensionName(file.Path)) = "xlsx" Then
            file.Delete
        End If
    Next file
End Sub


