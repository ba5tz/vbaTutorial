'Isi PlaceHolder Text
'-------------------------------------------------------
Public Sub PlaceHolder()
TBCustomer.Tag = "Isi Customer..."
TBPic.Tag = "Isi PIC..."
TBmanID.Tag = "Isi ID..."
TBManDate.Tag = "Isi Tanggal..."
End Sub

Public Sub PH_Enter(Ctrl As MSForms.TextBox)
If Ctrl.Value = Ctrl.Tag Then
    Ctrl.Value = ""
    Ctrl.ForeColor = vbWhite
End If
End Sub

Public Sub PH_Exit(Ctrl As MSForms.TextBox)
If Len(Ctrl.Value) = 0 Then
    Ctrl.Value = Ctrl.Tag
    Ctrl.ForeColor = &HB88A5F
End If
End Sub

'Penggunaan
'------------------------------------------

Private Sub TBCustomer_Enter()
PH_Enter TBCustomer
End Sub

Private Sub TBCustomer_Exit(ByVal Cancel As MSForms.ReturnBoolean)
PH_Exit TBCustomer
End Sub
