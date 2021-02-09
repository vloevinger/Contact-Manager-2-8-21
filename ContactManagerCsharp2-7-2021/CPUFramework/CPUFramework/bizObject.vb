Public Class bizObject(Of T)

    Dim sTableName As String
    Dim SPrimaryKeyName As String
    Dim SPrimaryKeyParameterName As String = ""
    Dim sGetSproc As String
    Dim sUpdateSproc As String
    Dim sDeleteSproc As String
    Dim objPrimaryTable As DataTable
    Dim objAdapter As SqlDataAdapter = New SqlDataAdapter


    Public Sub New(TableName As String)
        If TableName <> "" Then
            sTableName = TableName
            SPrimaryKeyName = "i" & sTableName & "Id"
            SPrimaryKeyParameterName = "@" & SPrimaryKeyName
            sGetSproc = sTableName & "Get"
            sUpdateSproc = sTableName & "Update"
            sDeleteSproc = sTableName & "Delete"
            Me.SetUpAdapter()
        End If
    End Sub

    Private Sub SetUpAdapter()
        With objAdapter
            .SelectCommand = mdUtility.GetSQLCommand(sGetSproc)
            .UpdateCommand = mdUtility.GetSQLCommand(sUpdateSproc)
            .InsertCommand = mdUtility.GetSQLCommand(sUpdateSproc)
            .DeleteCommand = mdUtility.GetSQLCommand(sDeleteSproc)
        End With
    End Sub
    Public Function Load(PrimaryKeyValue As Long) As DataTable
        objAdapter.SelectCommand.Parameters(SPrimaryKeyParameterName).Value = PrimaryKeyValue
        objPrimaryTable = mdUtility.ExecuteSQL(objAdapter)
        Return objPrimaryTable
    End Function

    Public Function GetList(Optional IncludeBlank = False, Optional IncludeAll = False) As DataTable

        With objAdapter.SelectCommand
            If .Parameters.Contains("@bAll") = False Then
                Throw New NotImplementedException
            End If
            .Parameters(SPrimaryKeyParameterName).Value = 0
            .Parameters("@bAll").Value = 1
            If IncludeBlank = True Then
                If .Parameters.Contains("@bIncludeBlank") = False Then
                    Throw New CPUException(.CommandText & " paramaters does not contain bIncludeBlank")
                End If
                .Parameters("@bIncludeBlank").Value = True
            End If
            If IncludeAll = True Then
                If .Parameters.Contains("@bIncludeAll") = False Then
                    Throw New CPUException(.CommandText & " paramaters does not contain bIncludeAll")
                End If
                .Parameters("@bIncludeAll").Value = True
            End If
        End With
        Dim objTable As DataTable = mdUtility.ExecuteSQL(objAdapter)
        Return objTable
    End Function
    Public Sub Save()
        If objPrimaryTable Is Nothing = False Then
            mdUtility.ExecuteSqlForUpdateOrDelete(objAdapter, objPrimaryTable)
        End If
    End Sub
    Public Sub CreateNew()
        If objPrimaryTable Is Nothing = True OrElse objPrimaryTable.Columns.Count = 0 Then
            Me.Load(0)
        End If

        objPrimaryTable.Rows.Clear()
        objPrimaryTable.Rows.Add()

    End Sub
    Public Sub Delete()
        With objPrimaryTable
            If .Rows.Count > 0 Then
                .Rows(0).Delete()
                mdUtility.ExecuteSqlForUpdateOrDelete(objAdapter, PrimaryTableObject)
            End If
        End With

    End Sub

    Public ReadOnly Property PrimaryTableObject As DataTable
        Get
            If objPrimaryTable Is Nothing Then
                Me.CreateNew()
            End If
            Return objPrimaryTable
        End Get
    End Property

    Private Function IsRowValidForRead(RowCollectionObj As DataRowCollection, FieldName As String, ForRead As Boolean) As Boolean
        Dim b As Boolean = False
        If IsNumeric(FieldName) Then
            Throw New Exception("Error trying to read value from table because field name is numeric.")
        End If
        With RowCollectionObj
            If .Count > 0 AndAlso (IsDBNull(.Item(0).Item(FieldName)) = False Or ForRead = False) Then
                b = True
            End If
        End With
        Return b
    End Function

    Public Property GetPrimaryTableFieldValueAsString(FieldName As T) As String
        Get
            Dim sValue As String = ""
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, True) = True Then
                    sValue = .Rows(0).Item(FieldName.ToString)
                End If
            End With
            Return sValue
        End Get
        Set(value As String)
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, False) = True Then
                    .Rows(0).Item(FieldName.ToString) = value
                End If
            End With

        End Set
    End Property
    Public Property GetPrimaryTableFieldValueAsInteger(FieldName As T) As Integer
        Get
            Dim nValue As Integer = 0
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, True) = True Then
                    nValue = .Rows(0).Item(FieldName.ToString)
                End If
            End With
            Return nValue
        End Get
        Set(value As Integer)
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, False) = True Then
                    .Rows(0).Item(FieldName.ToString) = value
                End If
            End With
        End Set
    End Property
    Public Property GetPrimaryTableFieldValueAsDecimal(FieldName As T) As Decimal
        Get
            Dim nValue As Decimal = 0
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, True) = True Then
                    nValue = .Rows(0).Item(FieldName.ToString)
                End If
            End With
            Return nValue
        End Get
        Set(value As Decimal)
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, False) = True Then
                    .Rows(0).Item(FieldName.ToString) = value
                End If
            End With
        End Set
    End Property
    Public Property GetPrimaryTableFieldValueAsLong(FieldName As T) As Long
        Get
            Dim nValue As Long = 0
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, True) = True Then
                    nValue = .Rows(0).Item(FieldName.ToString)
                End If
            End With
            Return nValue
        End Get
        Set(value As Long)
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, False) = True Then
                    .Rows(0).Item(FieldName.ToString) = value
                End If
            End With

        End Set
    End Property
    Public Property GetPrimaryTableFieldValueAsBoolean(FieldName As T) As Boolean
        Get
            Dim bValue As Boolean = False
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, True) = True Then
                    bValue = .Rows(0).Item(FieldName.ToString)
                End If
            End With
            Return bValue
        End Get
        Set(value As Boolean)
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, False) = True Then
                    .Rows(0).Item(FieldName.ToString) = value
                End If
            End With
        End Set
    End Property
    Public Property GetPrimaryTableFieldValueAsDate(FieldName As T) As Date
        Get
            Dim dValue As Date
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, True) = True Then
                    dValue = .Rows(0).Item(FieldName.ToString)
                End If
            End With
            Return dValue
        End Get
        Set(value As Date)
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, FieldName.ToString, False) = True Then
                    .Rows(0).Item(FieldName.ToString) = value
                End If
            End With

        End Set
    End Property

    Public ReadOnly Property PrimaryKeyName() As String
        Get
            Return SPrimaryKeyName
        End Get
    End Property
    Public Property PrimaryKeyValue() As Long
        Get
            Dim nId As Long = 0
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, SPrimaryKeyName, True) = True Then
                    nId = .Rows(0).Item(SPrimaryKeyName)
                End If
            End With
            Return nId
        End Get
        Set(value As Long)
            With Me.PrimaryTableObject
                If Me.IsRowValidForRead(.Rows, SPrimaryKeyName, False) = True Then
                    .Rows(0).Item(SPrimaryKeyName) = value
                End If
            End With
        End Set
    End Property

    Public Function FieldName(FieldNameValue As T) As String
        Dim sFieldName As String = FieldNameValue.ToString
        Return sFieldName
    End Function

    Public Sub SetCommandParamValue(CommandObj As SqlCommand, ParamaterName As T, Value As Object)
        Dim sParamName As String = "@" & ParamaterName.ToString
        With CommandObj.Parameters
            If .Contains(sParamName) = False Then
                Throw New CPUException(sParamName & "does not exist in " & CommandObj.CommandText)
            End If
            .Item(sParamName).Value = Value
        End With
    End Sub
End Class
