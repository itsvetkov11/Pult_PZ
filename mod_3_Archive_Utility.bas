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
    ws.Protect
    
    Application.StatusBar = False
    Application.EnableEvents = True
    
    ' Возвращаем курсор в поле поиска ЗВР
    ws.Range("PZ_SearchZVR").Select
End Sub

Sub Update_Bases_Manual()
    Dim msg As String
    
    msg = "Согласно принципа разумной достаточности, человека мы пока заменять не будем! :)" & vbCrLf & vbCrLf & _
          "Пожалуйста, обновляйте базы штатным способом (как показано на картинке чуть ниже уведомления):" & vbCrLf & _
          "Вкладка 'Данные' -> 'Обновить всё'." & vbCrLf & vbCrLf & _
          "Это самый надежный вариант для работы в общей сети."
          
    MsgBox msg, vbInformation, "РМЦ: Инструкция по обновлению"
End Sub

' ГЛАВНЫЙ ДИСПЕТЧЕР БЭКАПОВ
Sub Run_Smart_Backup_Logic()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PZ_Control")
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("Settings")
    
    ' 1. Бронежилет: Если Пульт открыт только для чтения, мы не сможем обновить даты
    If ThisWorkbook.ReadOnly Then Exit Sub
    
    Dim wbName As String: wbName = Trim(wsP.Range("PZ_DBName").Text)
    Dim lastAM As Date: lastAM = wsS.Range("Last_AM_Backup").Value
    Dim last11 As Date: last11 = wsS.Range("Last_11_Backup").Value
    Dim currDate As Date: currDate = Date
    Dim currHour As Integer: currHour = Hour(Now)
    
    Dim needBackup As Boolean: needBackup = False
    Dim backupType As String: backupType = ""

    ' 2. Логика "Кто первый встал" (Утренний бэкап)
    If lastAM < currDate Then
        needBackup = True: backupType = "AM": wsS.Range("Last_AM_Backup").Value = currDate
    
    ' 3. Логика "11-часовой чекпоинт"
    ElseIf currHour >= 11 And last11 < currDate Then
        needBackup = True: backupType = "11AM": wsS.Range("Last_11_Backup").Value = currDate
    End If

    ' 4. Выполнение, если сработал триггер
    If needBackup Then
        Execute_Silent_Backup wbName, backupType
        Clean_Old_Backups ' Запускаем санитара (7 дней)
    End If
End Sub

' ТИХАЯ КОПИЯ ФАЙЛА
Private Sub Execute_Silent_Backup(wbName As String, bType As String)
    On Error Resume Next
    Dim wbBase As Workbook: Set wbBase = Workbooks(wbName)
    If wbBase Is Nothing Then Exit Sub ' База не открыта
    
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
    
    Application.StatusBar = "MES: Создан резервный слепок базы (" & bType & ")"
End Sub

' САНИТАР (Удаление старых файлов > 7 дней)
Private Sub Clean_Old_Backups()
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(ThisWorkbook.Path & "\_MES_Backups\")
    Dim file As Object
    
    For Each file In folder.Files
        If DateDiff("d", file.DateCreated, Now) > 7 Then
            file.Delete
        End If
    Next file
End Sub

