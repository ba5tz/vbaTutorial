Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
Dim CmdBtn As CommandBarButton, CmdBtn2 As CommandBarButton
Dim CMenu As CommandBar
Dim Menu2 As CommandBarControl

On Error Resume Next
Set CMenu = Application.CommandBars("cell")
With CMenu
    .Controls("Hari ini").Delete
    .Controls("Menu Dua").Delete
    Set CmdBtn = .Controls.Add(temporary:=True, before:=1)
    Set CmdBtn2 = .Controls.Add(temporary:=True, before:=3)
End With
With CmdBtn
    .Caption = "Hari ini"
    .FaceId = "125"
    .OnAction = "menu_saya"
End With
With CmdBtn2
    .Caption = "Subscribe"
    .FaceId = "2083"
    .OnAction = "SubScribe"
End With

Set Menu2 = CMenu.Controls.Add(Type:=msoControlPopup, temporary:=True, before:=2)
With Menu2
    .Caption = "Menu Dua"
    With .Controls.Add
        .Caption = "ini Sub Menu"
        .FaceId = "23"
        .OnAction = "KlikMenu1"
    End With
    With .Controls.Add
        .Caption = "ini Sub Menu 2"
        .FaceId = "40"
         .OnAction = "KlikMenu2"
    End With
End With
End Sub
      
'------------------------------------  
'Simpan di Module
'------------------------------------      
      
      Public Sub menu_saya()
        ActiveCell.Value = Date
      End Sub

      Public Sub SubScribe()
        ThisWorkbook.FollowHyperlink ("https://www.youtube.com/c/AndiSetiadii?sub_confirmation=1")
      End Sub
      
      Public Sub KlikMenu1()
        msgbox "Sub menu 1 di klik"
      End Sub

      Public Sub KlikMenu2()
        msgbox "Sub menu 2 di klik"
      End Sub

