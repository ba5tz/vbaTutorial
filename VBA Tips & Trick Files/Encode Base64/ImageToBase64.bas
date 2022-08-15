' Script ini membutuhkan referensi ADO dan XML
' Tutorial lengkapnya bisa dilihat di https://youtu.be/dDCF8OIh-lk

Sub TestBase64()
    Dim bytes, Hasil
    With CreateObject("ADODB.Stream")
        .Open
        .Type = ADODB.adTypeBinary
        .LoadFromFile "C:\Users\andi\OneDrive\Gambar\ExcelKita\screenshot.png"
        bytes = .Read
        .Close
    End With
    Hasil = EncodeBase64(bytes)

    Sheet1.Range("C2").Value = Hasil

End Sub

Private Function EncodeBase64(bytes) As String

    Dim objXML    As MSXML2.DOMDocument
    Dim objNode   As MSXML2.IXMLDOMElement


    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = bytes
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing
End Function
