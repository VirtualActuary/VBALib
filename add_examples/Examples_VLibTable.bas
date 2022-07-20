Attribute VB_Name = "Examples_VLibTable"
Option Explicit

Function Example_table_operations()
    
    Dim abspath As String
    abspath = ThisWorkbook.Path & "\tests\TestTable.xlsx"
    
    Dim fso As FileSystemObject

    VLib.MkDirRecursive VLib.GetDirectoryName(abspath)
    
    Dim WB As Workbook
    Set WB = Fn.ExcelBook(abspath)
    
    Dim col As Collection
    Set col = Fn.col(Fn.dict("col1", "a", "col2", "b", "col3", "c"), Fn.dict("col1", 1, "col2", 2, "col3", 3))
                     
    Dim TestTable As ListObject
    Set TestTable = Fn.DictsToTable(col, WB.Worksheets(1).Range("A1"), "Table1")

    Dim TableObject As z_VLibTable
    Set TableObject = New z_VLibTable
    TableObject.Initialize TestTable
    
    Debug.Print TableObject.DataRowCount()  ' 2
    Debug.Print TableObject.Name()  ' Table1
    Debug.Print TableObject.CellValue(1, 2) ' b
    TableObject.CellValue(1, 2) = 3
    
    Debug.Print TableObject.CellValue(1, 2)  ' 3
    Debug.Print TableObject.ColumnRange(2).count()  ' 2
    
    TableObject.Resize (4)
    Debug.Print TableObject.DataRowCount()  ' 4
    
    WB.Close False
    
End Function
