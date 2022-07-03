Private Sub UserForm_Initialize()
  Rem Author : Andi Setiadi
Dim Datanya As Range
Dim Baris As Long
Dim Kamus As Object

Baris = Sheet1.Range("D" & Rows.Count).End(xlUp).Row
Set Kamus = CreateObject("Scripting.Dictionary")
Set Datanya = Range("D10:D" & Baris)

For Each isi In Datanya
    If Not Kamus.exists(LCase(isi.Value)) Then
        ComboBox1.AddItem isi.Value
        Kamus.Add LCase(isi.Value), isi.Value
    End If
Next isi
End Sub
