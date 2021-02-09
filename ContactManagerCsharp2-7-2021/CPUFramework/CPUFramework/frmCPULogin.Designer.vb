<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCPULogin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tblMain = New System.Windows.Forms.TableLayoutPanel()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.lblDBName = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.tblMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'tblMain
        '
        Me.tblMain.AutoSize = True
        Me.tblMain.ColumnCount = 3
        Me.tblMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.tblMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tblMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.tblMain.Controls.Add(Me.lblDBName, 0, 1)
        Me.tblMain.Controls.Add(Me.txtServer, 1, 0)
        Me.tblMain.Controls.Add(Me.txtDatabase, 1, 1)
        Me.tblMain.Controls.Add(Me.txtUsername, 1, 2)
        Me.tblMain.Controls.Add(Me.txtPassword, 1, 3)
        Me.tblMain.Controls.Add(Me.lblUsername, 0, 2)
        Me.tblMain.Controls.Add(Me.lblPassword, 0, 3)
        Me.tblMain.Controls.Add(Me.lblServer, 0, 0)
        Me.tblMain.Controls.Add(Me.btnCancel, 2, 4)
        Me.tblMain.Controls.Add(Me.btnOk, 1, 4)
        Me.tblMain.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tblMain.Location = New System.Drawing.Point(3, 1)
        Me.tblMain.Name = "tblMain"
        Me.tblMain.RowCount = 5
        Me.tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tblMain.Size = New System.Drawing.Size(347, 149)
        Me.tblMain.TabIndex = 0
        '
        'lblServer
        '
        Me.lblServer.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblServer.AutoSize = True
        Me.lblServer.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblServer.Location = New System.Drawing.Point(3, 6)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(91, 17)
        Me.lblServer.TabIndex = 0
        Me.lblServer.Text = "Server Name"
        '
        'lblDBName
        '
        Me.lblDBName.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblDBName.AutoSize = True
        Me.lblDBName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBName.Location = New System.Drawing.Point(3, 35)
        Me.lblDBName.Name = "lblDBName"
        Me.lblDBName.Size = New System.Drawing.Size(69, 17)
        Me.lblDBName.TabIndex = 1
        Me.lblDBName.Text = "Database"
        '
        'txtServer
        '
        Me.tblMain.SetColumnSpan(Me.txtServer, 2)
        Me.txtServer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtServer.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServer.Location = New System.Drawing.Point(100, 3)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(244, 23)
        Me.txtServer.TabIndex = 2
        '
        'txtDatabase
        '
        Me.tblMain.SetColumnSpan(Me.txtDatabase, 2)
        Me.txtDatabase.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatabase.Location = New System.Drawing.Point(100, 32)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(244, 23)
        Me.txtDatabase.TabIndex = 3
        '
        'txtUsername
        '
        Me.tblMain.SetColumnSpan(Me.txtUsername, 2)
        Me.txtUsername.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtUsername.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsername.Location = New System.Drawing.Point(100, 61)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(244, 23)
        Me.txtUsername.TabIndex = 4
        '
        'txtPassword
        '
        Me.tblMain.SetColumnSpan(Me.txtPassword, 2)
        Me.txtPassword.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.Location = New System.Drawing.Point(100, 90)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(244, 23)
        Me.txtPassword.TabIndex = 5
        '
        'lblUsername
        '
        Me.lblUsername.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUsername.Location = New System.Drawing.Point(3, 64)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(79, 17)
        Me.lblUsername.TabIndex = 6
        Me.lblUsername.Text = "User Name"
        '
        'lblPassword
        '
        Me.lblPassword.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassword.Location = New System.Drawing.Point(3, 93)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(69, 17)
        Me.lblPassword.TabIndex = 7
        Me.lblPassword.Text = "Password"
        '
        'btnOk
        '
        Me.btnOk.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnOk.AutoSize = True
        Me.btnOk.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.Location = New System.Drawing.Point(202, 119)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 27)
        Me.btnOk.TabIndex = 8
        Me.btnOk.Text = "Login"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.AutoSize = True
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(283, 119)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(61, 27)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmCPULogin
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(357, 159)
        Me.Controls.Add(Me.tblMain)
        Me.Name = "frmCPULogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.tblMain.ResumeLayout(False)
        Me.tblMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tblMain As Windows.Forms.TableLayoutPanel
    Friend WithEvents lblDBName As Windows.Forms.Label
    Friend WithEvents txtServer As Windows.Forms.TextBox
    Friend WithEvents txtDatabase As Windows.Forms.TextBox
    Friend WithEvents txtUsername As Windows.Forms.TextBox
    Friend WithEvents txtPassword As Windows.Forms.TextBox
    Friend WithEvents lblUsername As Windows.Forms.Label
    Friend WithEvents lblPassword As Windows.Forms.Label
    Friend WithEvents btnOk As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents lblServer As Windows.Forms.Label
End Class
