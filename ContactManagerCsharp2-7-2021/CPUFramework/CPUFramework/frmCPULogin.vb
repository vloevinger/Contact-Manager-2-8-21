Public Class frmCPULogin
    Dim bSuccess As Boolean = False

    Public Function ShowForm(ServerName As String, DatabaseName As String, Username As String, Password As String) As Boolean
        txtServer.Text = ServerName
        txtDatabase.Text = DatabaseName
        txtUsername.Text = Username
        txtPassword.Text = Password
        Me.ShowDialog()
        Return bSuccess
    End Function
    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Me.doLogin()
        If bSuccess = True Then
            Me.Hide()
        End If
    End Sub
    Private Sub doLogin()
        If txtServer.Text > "" And txtDatabase.Text > "" Then
            Try
                mdUtility.login(txtServer.Text, txtDatabase.Text, txtUsername.Text, txtPassword.Text, True)
                bSuccess = True
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Else
            MsgBox("Server and Database cannot be blank", vbOKOnly)
        End If
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        bSuccess = False
        Me.Hide()
    End Sub

    Public ReadOnly Property ServerName As String
        Get
            Return txtServer.Text
        End Get
    End Property

    Public ReadOnly Property DatabaseName As String
        Get
            Return txtDatabase.Text
        End Get
    End Property

    Public ReadOnly Property Username As String
        Get
            Return txtUsername.Text
        End Get
    End Property

    Public ReadOnly Property Password As String
        Get
            Return txtPassword.Text
        End Get
    End Property
End Class