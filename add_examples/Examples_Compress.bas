Attribute VB_Name = "Examples_Compress"
Option Explicit


Public Sub Example_zip()

    Dim ZipObj As z__Compress
    Set ZipObj = New z__Compress

    Dim abspath As String
    abspath = ThisWorkbook.Path & "\tests\writeToCSVTest.csv"

    CreateCSVFile abspath

    ZipObj.zip (abspath)
    
End Sub


Public Sub Example_UnZip()
    Example_zip

    Dim ZipObj As z__Compress
    Set ZipObj = New z__Compress
    
    ZipObj.UnZip (ThisWorkbook.Path & "\tests\writeToCSVTest.zip")
    
End Sub


Private Sub CreateCSVFile(abspath As String)
    VLib.MkDirRecursive VLib.GetDirectoryName(abspath)
    
    ' Create the csv object
    Dim csvout As z_WsCsvInterface
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

    ' Write to selected csv file.
    With csvout
        .parseConfig.Path = abspath
        .ExportToCSV arr
        If .exportSuccess = False Then
            Err.Raise 6, , "CSV write to " & abspath & " failed, check file permissions or if open in another application."
        End If
    End With
End Sub
