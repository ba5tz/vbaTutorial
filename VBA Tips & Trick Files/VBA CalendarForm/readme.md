# Calendar Form

## File:
<a href="https://github.com/ba5tz/vbaTutorial/blob/master/VBA%20Tips%20%26%20Trick%20Files/VBA%20CalendarForm/CalendarForm.rar">Download CalendarForm.rar</a>

## Worksheet Script:
```vb

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim Tanggal As Date

If Not Intersect(Target, Range("C2:C20")) Is Nothing Then
    Tanggal = CalendarForm.GetDate
    If Not Tanggal = Empty Then
        Target = Tanggal
    End If
End If
End Sub
```

## Function Script:
```vb
Function GetTanggal(Ctr as Control)
Dim Tanggal As Date

Tanggal = CalendarForm.GetDate
If Not Tanggal = Empty Then
    Ctr.Text = Tanggal
End If
End Sub
```

## Penggunaan:
```vb
GetTanggal Textbox1.Text
```
