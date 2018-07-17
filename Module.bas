Public Sub WriteToFile(FilePath As String, ContentToWrite As String)
    Dim utfStr() As Byte
    utfStr = ADO_EncodeUTF8(ContentToWrite)
    Open FilePath For Binary Lock Read Write As #1
    Seek #1, LOF(1) + 1
    Put #1, , utfStr    
    Close #1
End Sub

Public Function ADO_EncodeUTF8(ByVal strUTF16 As String) As Byte() 
    Dim objStream As Object
    Dim data() As Byte    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeText
    objStream.Open
    objStream.WriteText strUTF16
    objStream.flush
    objStream.Position = 0
    objStream.Type = adTypeBinary
    objStream.Read 3 ' skip BOM
    data = objStream.Read()
    objStream.Close
    ADO_EncodeUTF8 = data    
End Function
