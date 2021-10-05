Attribute VB_Name = "Module1"
'Copyright Ellis Clarke 2021, This code is under MIT Licence'
Sub Convert()

Dim excelRange As Range
Dim jsonItems As New Collection
Dim jsonDictionary As New Dictionary
Dim jsonFileObject As New FileSystemObject
Dim jsonFileExport As TextStream
Dim i As Long
Dim cell As Variant

Set excelRange = Cells(1, 1).CurrentRegion

For i = 2 To excelRange.Rows.Count
    jsonDictionary("name") = Cells(i, 1)
    jsonDictionary("orbital") = Cells(i, 2)
    jsonDictionary("colour") = Cells(i, 3)


    jsonItems.Add jsonDictionary
    Set jsonDictionary = Nothing
Next i


Set jsonFileExport = jsonFileObject.CreateTextFile("INSERT DIRECTORY HERE", True)
jsonFileExport.WriteLine (JsonConverter.ConvertToJson(jsonItems, Whitespace:=3))


End Sub
