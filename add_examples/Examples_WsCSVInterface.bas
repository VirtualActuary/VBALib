Attribute VB_Name = "Examples_WsCSVInterface"
Option Explicit


Function Example_ExportToCSV()
    Dim fso As New FileSystemObject
    Dim csvout As z_WsCsvInterface
    Dim abspath As String
    abspath = ThisWorkbook.Path & "\tests\writeToCSVTest.csv"
    
    ' Create the csv object
    Set csvout = New z_WsCsvInterface
    With csvout.parseConfig
        .dialect.fieldsDelimiter = ","
        .dialect.recordsDelimiter = vbCrLf
        .Headers = True
    End With
    
    Dim arr(1, 1) As Variant
    arr(0, 0) = "a"
    arr(0, 1) = "b"
    arr(1, 0) = "c"
    arr(1, 1) = "d"
    
    ' Ensure the directory exists
    VLib.MkDirRecursive VLib.GetDirectoryName(abspath)
    
    ' Write to selected csv file.
    With csvout
        .parseConfig.Path = abspath
        .ExportToCSV arr
        If .exportSuccess = False Then
            Err.Raise 6, , "CSV write to " & abspath & " failed, check file permissions or if open in another application."
        End If
    End With
    
End Function


Function Example_GetDataFromCSV()
    Dim CSVIn As z_WsCsvInterface
    Set CSVIn = New z_WsCsvInterface
    Debug.Print CSVIn.GetDataFromCSV(ThisWorkbook.Path & "\tests\writeToCSVTest.csv")
End Function


