Private Sub UserForm_Initialize()
    cmbEnv.AddItem "DEV"
    cmbEnv.AddItem "UAT"
    cmbEnv.AddItem "PROD"
    cmbEnv.ListIndex = 0
End Sub

Private Sub btnLogin_Click()
    If txtUser.Text = "" Or txtPwd.Text = "" Then
        MsgBox "Enter username and password."
        Exit Sub
    End If
    Me.Hide
End Sub