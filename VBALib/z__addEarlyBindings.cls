VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "z__addEarlyBindings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Add references for this project programatically. If you are uncertain what to put here,
' Go to Tools -> References and use the filename of the reference (eg. msado15.dll for
' Microsoft ActiveX Data Objects 6.1 Library'), then run getPackageGUID("msado15.dll")
' to see what options you have:
'**********************************************************************************
'* Add selected references to this project
'**********************************************************************************
Public Sub addEarlyBindings()
    On Error GoTo ErrorHandler
        If Not isBindingNameLoaded("ADODB") Then
            'Microsoft ActiveX Data Objects 6.1
            ThisWorkbook.VBProject.References.addFromGuid "{B691E011-1797-432E-907A-4D8C69339129}", 6.1, 0
        End If
        
        If Not isBindingNameLoaded("VBIDE") Then
            'Microsoft Visual Basic for Applications Extensibility 5.3
            ThisWorkbook.VBProject.References.addFromGuid "{0002E157-0000-0000-C000-000000000046}", 5.3, 0
        End If
        
        
        If Not isBindingNameLoaded("Scripting") Then
            'Microsoft Scripting Runtime version 1.0
            ThisWorkbook.VBProject.References.addFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
        End If
        
    
        If Not isBindingNameLoaded("VBScript_RegExp_55") Then
            'Microsoft VBScript Regular Expressions 5.5
            ThisWorkbook.VBProject.References.addFromGuid "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5
        End If
        
        If Not isBindingNameLoaded("Shell32") Then
            'Microsoft Shell Controls And Automation
            ThisWorkbook.VBProject.References.addFromGuid "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}", 1, 0
        End If
    
    Exit Sub
ErrorHandler:
End Sub


'**********************************************************************************
'* Verify if a reference is loaded
'**********************************************************************************
Public Function isBindingNameLoaded(ref As String) As Boolean
    ' https://www.ozgrid.com/forum/index.php?thread/62123-check-if-ref-library-is-loaded/&postID=575116#post575116
    isBindingNameLoaded = False
    Dim xRef As Variant
    For Each xRef In ThisWorkbook.VBProject.References
        If VBA.LCase(xRef.Name) = VBA.LCase(ref) Then
            isBindingNameLoaded = True
        End If
    Next xRef
    
End Function


'**********************************************************************************
'* Print all current active GUIDs
'**********************************************************************************
Private Sub printAllEarlyBindings()
    ' https://www.ozgrid.com/forum/index.php?thread/62123-check-if-ref-library-is-loaded/&postID=575116#post575116
    Dim xRef As Variant
    For Each xRef In ThisWorkbook.VBProject.References
        Debug.Print "**************" & xRef.Name
        Debug.Print xRef.Description
        Debug.Print xRef.Major
        Debug.Print xRef.Minor
        Debug.Print xRef.FullPath
        Debug.Print xRef.GUID
        Debug.Print vbNullString
    Next xRef
    
End Sub

