Public Class bizUserLogin
    Inherits bizObject(Of FieldEnum)

    Public Enum FieldEnum
        vchUserName
        vchPassword
        iRole
        vchFirstName
        vchLastName
        vchEmail
    End Enum

    Public Sub New()
        MyBase.New("UserLogin")
    End Sub

    Public Sub LoadByUsernamePassword()
        Dim objcmd As SqlCommand = mdUtility.GetSQLCommand("UserLoginGet")
        Me.SetCommandParamValue(objcmd, FieldEnum.vchUserName, Me.UserName)
        Me.SetCommandParamValue(objcmd, FieldEnum.vchPassword, Me.Password)

        Dim objT As DataTable = mdUtility.ExecuteSQL(objcmd)
        If objT.Rows.Count = 0 Then
            Throw New CPUException("Invalid Login")
        End If

        Me.Load(objT.Rows(0).Item("iUserLoginId"))

    End Sub

    Public Property UserName As String
        Get
            Return Me.GetPrimaryTableFieldValueAsString(FieldEnum.vchUserName)
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(FieldEnum.vchUserName.ToString) = value
        End Set
    End Property

    Public Property Password As String
        Get
            Return Me.GetPrimaryTableFieldValueAsString(FieldEnum.vchPassword)
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(FieldEnum.vchPassword.ToString) = value
        End Set
    End Property

    Public Property Role As Integer
        Get
            Return Me.GetPrimaryTableFieldValueAsInteger(FieldEnum.iRole)
        End Get
        Set(value As Integer)
            Me.PrimaryTableObject.Rows(0).Item(FieldEnum.iRole.ToString) = value
        End Set
    End Property

    Public Property FirstName As String
        Get
            Return Me.GetPrimaryTableFieldValueAsString(FieldEnum.vchFirstName)
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(FieldEnum.vchFirstName.ToString) = value
        End Set
    End Property

    Public Property LastName As String
        Get
            Return Me.GetPrimaryTableFieldValueAsString(FieldEnum.vchLastName)
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(FieldEnum.vchLastName.ToString) = value
        End Set
    End Property

    Public Property Email As String
        Get
            Return Me.GetPrimaryTableFieldValueAsString(FieldEnum.vchEmail)
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(FieldEnum.vchEmail.ToString) = value
        End Set
    End Property


End Class
