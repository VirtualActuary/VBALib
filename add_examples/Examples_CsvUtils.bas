Attribute VB_Name = "Examples_CsvUtils"
Option Explicit
' Most CSV functions here aren't currently being used and will possibly not be used in
' the future. These examples are provided for incase someone might find it useful.


Function Example_ParseCSVToCollection()
    ' ParseCSVToCollection() returns a Collection of records, and the record is a
    ' collection of fields. If error occurs, it returns Nothing and the error
    ' information is set in Err object. Optional boolean argument allowVariableNumOfFields
    ' specifies whether variable number of fields in records is allowed or handled as error.

    Dim csv As Collection
    Dim rec As Collection, fld As Variant

    Set csv = CsvUtils.ParseCSVToCollection("aaa,bbb,ccc" & vbCr & "xxx,yyy,zzz")

    For Each rec In csv
      For Each fld In rec
        Debug.Print fld
      Next
    Next
    'Output:
    '   aaa
    '   bbb
    '   ccc
    '   xxx
    '   yyy
    '   zzz
    '   asdf
End Function


Function Example_ParseCSVToArray()
    ' ParseCSVToArray() returns a Variant that contains 2-dimensional array -
    ' String(1 To recordCount, 1 To fieldCount). If error occurs, it returns
    ' Null and the error information is set in Err object. If input text
    ' is zero-length (""), it returns empty array - String(0 To -1).
    ' Optional boolean argument allowVariableNumOfFields specifies whether
    ' variable number of fields in records is allowed or handled as error.
    
    Dim csv As Variant
    Dim i As Long, j As Variant

    csv = CsvUtils.ParseCSVToArray("aaa,bbb,ccc" & vbCr & "xxx,yyy,zzz")

    For i = LBound(csv, 1) To UBound(csv, 1)
      For j = LBound(csv, 2) To UBound(csv, 2)
        Debug.Print csv(i, j)
      Next
    Next
    ' Output:
    '    aaa
    '    bbb
    '    ccc
    '    xxx
    '    yyy
    '    zzz
End Function


Function Example_ConvertArrayToCSV()
    ' ConvertArrayToCSV() reads 2-dimensional array inArray and return CSV text.
    ' If error occurs, it return the string "", and the error information is set
    ' in Err object. fmtDate is used as the argument of text formatting function
    ' Format if an element of the array is Date type. The optional argument
    ' quoting specifies what type of fields to be quoted:
    '   MINIMAL: Quoting only if it is necessary (the field includes double-quotes, comma, line breaks).
    '   ALL: Quoting all the fields.
    '   NONNUMERIC: Similar to MINIMAL, but quoting also all the String type fields.
    ' The optional arugment recordSeparator specifies record separator (line terminator), default is CRLF.
    
    Dim csv As String
    Dim a(1 To 2, 1 To 2) As Variant
    a(1, 1) = DateSerial(1900, 4, 14)
    a(1, 2) = "Exposition Universelle de Paris 1900"
    a(2, 1) = DateSerial(1970, 3, 15)
    a(2, 2) = "Japan World Exposition, Osaka 1970"
    
    csv = CsvUtils.ConvertArrayToCSV(a, "yyyy/mm/dd")
    If Err.Number <> 0 Then
        Debug.Print Err.Number & " (" & Err.source & ") " & Err.Description
    End If
    
    Debug.Print csv
    ' Output:
    '    1900/04/14,Exposition Universelle de Paris 1900
    '    1970/03/15,"Japan World Exposition, Osaka 1970"
End Function


Function Example_ParseCSVToDictionary()
    ' ParseCSVToDictionary() returns a Dictionary (Scripting.Dictionary) of records;
    ' the records are Collections of fields. In default, the first field of each record
    ' is the key of the dictionary. The column number of the key field can be specified
    ' by keyColumn, whose default value is 1. If there are multiple records whose
    ' key fields are the same, the value for the key is set to the last record among them.
    ' If error occurs, it returns Nothing and the error information is set in Err object.
    ' Optional boolean argument allowVariableNumOfFields specifies whether variable number
    ' of fields in records is allowed or handled as error.
    
    Dim csv As String
    Dim csvd As Object

    csv = "key,val1, val2" & vbCrLf & "name1,v11,v12" & vbCrLf & "name2,v21,v22"
    Set csvd = CsvUtils.ParseCSVToDictionary(csv, 1)
    Debug.Print csvd("name1")(2)
    Debug.Print csvd("name1")(3)
    Debug.Print csvd("name2")(2)
    ' Output:
    '    v11
    '    v12
    '    v21
End Function


Function Example_GetFieldDictionary()
    ' GetFieldDictionary() returns a Dictionary (Scripting.Dictionary) of field names,
    ' whose keys are the field values of the first records and whose values are
    ' the column numbers of the fields. If there are multiple fields of the same value
    ' in the first record, the value for the key is set to the largest column number
    ' among the fields.
    ' If error occurs, it returns Nothing and the error information is set in Err object.
    
    Dim csv As String
    Dim csva
    Dim field As Object

    csv = "key,val1, val2" & vbCrLf & "name1,v11,v12" & vbCrLf & "name2,v21,v22"
    Set field = CsvUtils.GetFieldDictionary(csv)
    csva = CsvUtils.ParseCSVToArray(csv)
    Debug.Print csva(2, field("key"))
    Debug.Print csva(3, field("val1"))
    ' Output:
    '    name1
    '    v21
End Function
