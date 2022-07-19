Attribute VB_Name = "Examples_Compress"
Option Explicit


Function Example_zip()

    Dim ZipObj As z__Compress
    Set ZipObj = New z__Compress
    
    ZipObj.Zip (ThisWorkbook.Path & "\tests\writeToCSVTest.csv")
    
End Function


Function Example_UnZip()

    Dim ZipObj As z__Compress
    Set ZipObj = New z__Compress
    
    Debug.Print ZipObj.UnZip(ThisWorkbook.Path & "\tests\writeToCSVTest.zip")
    
End Function
