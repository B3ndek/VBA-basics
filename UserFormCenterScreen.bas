'"""
' & 'This module provides functionality to center a VBA UserForm on the screen.
' & '
' & 'Functions and Subroutines:
' & '  - Public Sub CenterUserForm(frm As Object)
' & '      Centers the provided UserForm in the middle of the screen. It adjusts the
' & '      UserForm position based on the screen dimensions and the UserForm size.
' & '
' & '  - Private Sub UserForm_Initialize()
' & '      Automatically centers the UserForm when it is initialized. Calls the
' & '      CenterUserForm procedure, passing the current UserForm (Me) as input.
' & '      This ensures that any UserForm using this subroutine will always
' & '      appear at the center of the screen when displayed.
' & '
' & 'Usage:
' & '  1) Ensure the module is included in your VBA Project.
' & '  2) Attach the UserForm_Initialize subroutine to the desired UserForm.
' & '  3) The UserForm will automatically center itself when opened.
"""

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
'