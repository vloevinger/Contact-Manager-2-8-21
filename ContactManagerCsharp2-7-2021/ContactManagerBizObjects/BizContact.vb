
Public Class BizContact
    Inherits bizObject(Of FieldEnum)

    Public Enum FieldEnum
        iContactId
        iAddressId
        vchFirstName
        vchLastName
        vchStreet
        vchCity
        chState
        vchPostalCode
    End Enum

    Public Sub New()
        MyBase.New("contact")
    End Sub
    Public Property AddressId As Long
        Get
            Dim nValue As Long = Me.GetPrimaryTableFieldValueAsLong(BizContact.FieldEnum.iAddressId)
            Return nValue
        End Get
        Set(value As Long)
            Me.PrimaryTableObject.Rows(0).Item(BizContact.FieldEnum.iAddressId.ToString) = value
        End Set
    End Property

    Public Property FirstName As String
        Get
            Dim sValue As String = Me.GetPrimaryTableFieldValueAsString(BizContact.FieldEnum.vchFirstName)
            Return sValue
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(BizContact.FieldEnum.vchFirstName.ToString) = value
        End Set
    End Property

    Public Property LasttName As String
        Get
            Dim sValue As String = Me.GetPrimaryTableFieldValueAsString(BizContact.FieldEnum.vchLastName)
            Return sValue
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(BizContact.FieldEnum.vchLastName.ToString) = value
        End Set
    End Property
    Public Property Street As String
        Get
            Dim sValue As String = Me.GetPrimaryTableFieldValueAsString(BizContact.FieldEnum.vchStreet)
            Return sValue
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(BizContact.FieldEnum.vchStreet.ToString) = value
        End Set
    End Property

    Public Property City As String
        Get
            Dim sValue As String = Me.GetPrimaryTableFieldValueAsString(BizContact.FieldEnum.vchCity)
            Return sValue
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(BizContact.FieldEnum.vchCity.ToString) = value
        End Set
    End Property

    Public Property State As String
        Get
            Dim sValue As String = Me.GetPrimaryTableFieldValueAsString(BizContact.FieldEnum.chState)
            Return sValue
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(BizContact.FieldEnum.chState.ToString) = value
        End Set
    End Property

    Public Property PostalCode As String
        Get
            Dim sValue As String = Me.GetPrimaryTableFieldValueAsString(BizContact.FieldEnum.vchPostalCode)
            Return sValue
        End Get
        Set(value As String)
            Me.PrimaryTableObject.Rows(0).Item(BizContact.FieldEnum.vchPostalCode.ToString) = value
        End Set
    End Property

    Public Function ListofContacts() As List(Of BizContact)
        Dim lst As List(Of BizContact) = New List(Of BizContact)
        Dim objT As DataTable = Me.GetList()

        For Each objR As DataRow In objT.Rows
            Dim nId As Long = objR.Item(Me.PrimaryKeyName)
            Dim objContact As BizContact = New BizContact
            objContact.Load(nId)
            lst.Add(objContact)
        Next

        Return lst
    End Function

    Public Function Search(AnyCriteria As String) As List(Of BizContact)
        Dim lst As List(Of BizContact) = New List(Of BizContact)
        Dim objT As DataTable
        Using objCmd As SqlCommand = mdUtility.GetSQLCommand("ContactSearch")
            objCmd.Parameters("@vchSearchCriteria").Value = AnyCriteria
            objT = mdUtility.ExecuteSQL(objCmd)
        End Using

        For Each objR As DataRow In objT.Rows
            Dim nId As Long = objR.Item(Me.PrimaryKeyName)
            Dim objContact As BizContact = New BizContact
            objContact.Load(nId)
            lst.Add(objContact)
        Next

        Return lst

    End Function

End Class
