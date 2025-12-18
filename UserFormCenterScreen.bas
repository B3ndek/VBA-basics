Public Sub CenterUserForm(frm As Object)
    With frm
        .StartUpPosition = 0
        .Left = Application.Left + (Application.Width - .Width) / 2
        .Top = Application.Top + (Application.Height - .Height) / 2
    End With
End Sub

Private Sub UserForm_Initialize()
    CenterUserForm Me
End Sub

