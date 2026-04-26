import glob

# reset find
for filepath in glob.glob('*.bas') + glob.glob('*.cls'):
    with open(filepath, 'r', encoding='utf-8' if 'mod_4' not in filepath and 'Лист' not in filepath else 'windows-1251') as f:
        content = f.read()

    # mod_2
    content = content.replace('    Set f = wsRef.Columns("G").Find(What:=deptName, LookIn:=xlValues, LookAt:=xlWhole)\n', '    Set f = wsRef.Columns("G").Find(What:=deptName, LookIn:=xlValues, LookAt:=xlWhole)\n    wsRef.Cells.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n')
    
    # mod_4
    if 'mod_4' in filepath:
        content = content.replace('        Set foundCell = wsSettings.Columns("H").Find(What:="Путь_ODM021", LookIn:=xlValues, LookAt:=xlWhole)\n', '        Set foundCell = wsSettings.Columns("H").Find(What:="Путь_ODM021", LookIn:=xlValues, LookAt:=xlWhole)\n        wsSettings.Cells.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n')
        
    # List1
    if 'Лист1' in filepath:
        content = content.replace('    Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(What:=sSearch, LookIn:=xlValues, LookAt:=xlWhole)\n', '    Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(What:=sSearch, LookIn:=xlValues, LookAt:=xlWhole)\n    tbl.DataBodyRange.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n')
        content = content.replace('        Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(What:=CLng(sSearch), LookIn:=xlValues, LookAt:=xlWhole)\n', '        Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(What:=CLng(sSearch), LookIn:=xlValues, LookAt:=xlWhole)\n        tbl.DataBodyRange.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n')
        content = content.replace('    Set foundRow = tbl.ListColumns(2).DataBodyRange.Find(What:=sSearch, LookIn:=xlValues, LookAt:=xlWhole)\n', '    Set foundRow = tbl.ListColumns(2).DataBodyRange.Find(What:=sSearch, LookIn:=xlValues, LookAt:=xlWhole)\n    tbl.DataBodyRange.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n')
        content = content.replace('        Set foundRow = tbl.ListColumns(2).DataBodyRange.Find(What:=CLng(sSearch), LookIn:=xlValues, LookAt:=xlWhole)\n', '        Set foundRow = tbl.ListColumns(2).DataBodyRange.Find(What:=CLng(sSearch), LookIn:=xlValues, LookAt:=xlWhole)\n        tbl.DataBodyRange.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n')

    # Ensure no duplicates
    content = content.replace('    wsRef.Cells.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n    wsRef.Cells.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n', '    wsRef.Cells.Find What:="", LookAt:=xlPart \' Сброс поиска (Ctrl+F) на частичное совпадение\n')

    with open(filepath, 'w', encoding='utf-8' if 'mod_4' not in filepath and 'Лист' not in filepath else 'windows-1251') as f:
        f.write(content)
