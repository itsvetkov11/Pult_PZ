import glob

# we need to append a dummy search with xlPart right after any xlWhole search to reset Excel's find state
def add_dummy_search(filename, encoding):
    with open(filename, 'r', encoding=encoding) as f:
        lines = f.readlines()
        
    new_lines = []
    for line in lines:
        new_lines.append(line)
        if "LookAt:=xlWhole" in line:
            # Add a dummy search to reset the dialog flag
            indent = line[:len(line) - len(line.lstrip())]
            new_lines.append(indent + "' Сброс поиска (Ctrl+F) на частичное совпадение\n")
            new_lines.append(indent + "Dim dummyRng As Range\n")
            # find what exactly was searched to know the range, or just use Cells.Find
            # Actually, the simplest reset is Cells.Find(What:="", LookAt:=xlPart)
            new_lines.append(indent + "Set dummyRng = Cells.Find(What:=\"\", LookAt:=xlPart)\n")
            
    with open(filename, 'w', encoding=encoding) as f:
        f.writelines(new_lines)

add_dummy_search('mod_2_Constructor.bas', 'utf-8')
add_dummy_search('mod_4_ODM021_Compare.bas', 'windows-1251')
add_dummy_search('Лист1.cls', 'windows-1251')
