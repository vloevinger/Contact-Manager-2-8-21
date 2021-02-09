Public Module mdUtility
    Dim SConnString As String = ""

    Public Sub Login(serverName As String, Database As String, UserName As String, Password As String, TryLogin As Boolean)
        SConnString = GetConnectionString(serverName, Database, UserName, Password)
        Dologin(SConnString, TryLogin)
    End Sub

    Public Sub Login(ConnectionStringValue As String, TryLogin As Boolean)
        Dologin(ConnectionStringValue, TryLogin)
    End Sub

    Private Sub Dologin(ConnectionStringValue As String, TryLogin As Boolean)
        SConnString = ConnectionStringValue
        'test whether user can log in
        If TryLogin = True Then
            Using objConn As SqlConnection = New SqlConnection
                With objConn
                    .ConnectionString = ConnectionStringValue
                    .Open()
                End With
            End Using
        End If
    End Sub

    Private Function GetConnectionString(serverName As String, Database As String, UserName As String, Password As String) As String
        Dim sConn As String = "Server=" & serverName & ";Database=" & Database & ";"
        If UserName = "" And Password = "" Then
            sConn = sConn & "Trusted_Connection = True;"
        Else
            sConn = sConn & "User ID = " & UserName & ";" & "Password= " & Password & ";"
        End If

        Return sConn
    End Function

    Private Sub DeriveParamenter(CommandObj As SqlCommand)
        Dim objConn As SqlConnection = New SqlConnection
        Using objConn
            With objConn
                .ConnectionString = SConnString
                'Debug.Print(.ConnectionString)
                .Open()
            End With
            CommandObj.Connection = objConn
            SqlCommandBuilder.DeriveParameters(CommandObj)
            objConn.Close()
        End Using
        For Each objP As SqlParameter In CommandObj.Parameters
            objP.Value = DBNull.Value
        Next
    End Sub

    Public Function GetSQLCommand(SprocName As String) As SqlCommand
        Dim ObjCmd As SqlCommand = New SqlCommand
        With ObjCmd
            .CommandType = CommandType.StoredProcedure
            .CommandText = SprocName
        End With
        DeriveParamenter(ObjCmd)
        Return ObjCmd
    End Function
    Public Sub ExecuteSqlForUpdateOrDelete(AdapterObj As SqlDataAdapter, TableObj As DataTable)
        With AdapterObj
            If .UpdateCommand Is Nothing = False Then
                SetupAdapterSourcers(.UpdateCommand, TableObj)
            End If
            If .InsertCommand Is Nothing = False Then
                SetupAdapterSourcers(.InsertCommand, TableObj)
            End If
            If .DeleteCommand Is Nothing = False Then
                SetupAdapterSourcers(.DeleteCommand, TableObj)
            End If
        End With
        DoExecuteSQL(AdapterObj, False, TableObj)
    End Sub

    Private Sub SetupAdapterSourcers(Commandobj As SqlCommand, Tableobj As DataTable)
        For Each objParam As SqlParameter In Commandobj.Parameters
            Dim sFieldName As String = Replace(objParam.ParameterName, "@", "")
            If Tableobj.Columns.Contains(sFieldName) Then
                objParam.SourceColumn = sFieldName
            End If
        Next

    End Sub
    Public Function ExecuteSQL(CommandObj As SqlCommand) As DataTable
        Dim objAdapter As SqlDataAdapter = New SqlDataAdapter
        objAdapter.SelectCommand = CommandObj
        Dim objT As DataTable = DoExecuteSQL(objAdapter, True, Nothing)
        Return objT
    End Function

    Public Function ExecuteSQL(AdaptorObj As SqlDataAdapter) As DataTable
        Dim objT As DataTable = DoExecuteSQL(AdaptorObj, True, Nothing)
        Return objT
    End Function


    Public Function ExecuteSQL(SQLStatement As String, Optional ReturnTable As Boolean = True) As DataTable
        Dim objCmd As SqlCommand = New SqlCommand(SQLStatement)
        Dim objAdapter As SqlDataAdapter = New SqlDataAdapter
        objAdapter.SelectCommand = objCmd
        Dim objT As DataTable = DoExecuteSQL(objAdapter, ReturnTable, Nothing)
        Return objT
    End Function


    Private Function DoExecuteSQL(AdaptorObj As SqlDataAdapter, ReturnTable As Boolean, TableObj As DataTable) As DataTable
        Dim objReturnTable As DataTable = New DataTable

        If SConnString = "" Then
            Throw New Exception("Connection string is blank. You must login.")
        End If

        Dim bModified As Boolean = False
        Dim bAdded As Boolean = False
        Dim bdeleted As Boolean = False

        If ReturnTable = False AndAlso TableObj Is Nothing = False Then
            bModified = IsDataRowStateChanged(AdaptorObj, TableObj, DataViewRowState.ModifiedCurrent)
            bAdded = IsDataRowStateChanged(AdaptorObj, TableObj, DataViewRowState.Added)
            bdeleted = IsDataRowStateChanged(AdaptorObj, TableObj, DataViewRowState.Deleted)
        End If

        Dim objConn As SqlConnection = New SqlConnection
        Using objConn
            With objConn
                .ConnectionString = SConnString
                'Debug.Print(.ConnectionString)
                .Open()
            End With
            SetAdapterConnection(AdaptorObj, objConn)
            With AdaptorObj
                Try
                    If ReturnTable = True Then
                        Dim objReader As SqlDataReader = .SelectCommand.ExecuteReader()
                        objReturnTable = New DataTable
                        objReturnTable.Load(objReader)
                        SetPropertiesForAllColumns(objReturnTable)
                    Else
                        ' .ExecuteNonQuery()
                        .Update(TableObj)
                    End If
                    Try
                        objConn.Close()
                    Catch ex As Exception
                    End Try
                Catch ex As Exception When ex.Message.ToLower.Contains("cannot insert the value null")
                    Throw New CPUException("Please fill out all fields.")
                    'Catch ex As Exception When ex.Message.ToLower.Contains("")
                    '    Throw New CPUException("Please fill out all fields.")
                Catch ex As Exception When ex.Message.ToLower.Contains("ck_")
                    Throw New CPUException(ParseConstraintViolation(ex.Message, "ck_"))
                Catch ex As Exception When ex.Message.ToLower.Contains("u_")
                    Throw New CPUException(ParseConstraintViolation(ex.Message, "u_"))
                Catch ex As Exception When ex.Message.ToLower.Contains("fk_")
                    Throw New CPUException(ParseConstraintViolation(ex.Message, "fk_"))
                Finally
                    WriteDebugSQL(AdaptorObj, TableObj, ReturnTable, bAdded, bModified, bdeleted)
                End Try
            End With

        End Using
        CheckReturnValueOfAdapterCommands(AdaptorObj, TableObj, ReturnTable, bAdded, bModified, bdeleted)
        Return objReturnTable
    End Function
    Private Sub SetPropertiesForAllColumns(Tableobj As DataTable)
        For Each objc As DataColumn In Tableobj.Columns
            With objc
                .AllowDBNull = True
                .ReadOnly = False
            End With
        Next
    End Sub
    Private Sub CheckReturnValueOfAdapterCommands(AdapterObj As SqlDataAdapter, TableObj As DataTable, ReturnTable As Boolean, RowsAdded As Boolean, RowsModified As Boolean, RowsDeleted As Boolean)
        Dim nReturnValue As Integer = 0
        Dim sMessage As String = ""
        With AdapterObj
            If ReturnTable = True Then
                If .SelectCommand Is Nothing = False Then
                    nReturnValue = GetReturnValue(.SelectCommand)
                    sMessage = GetSprocMessage(.SelectCommand, True)
                End If

            Else
                If RowsAdded = True Then
                    nReturnValue = GetReturnValue(.InsertCommand)
                    sMessage = GetSprocMessage(.InsertCommand, True)
                ElseIf RowsModified = True Then
                    nReturnValue = GetReturnValue(.UpdateCommand)
                    sMessage = GetSprocMessage(.UpdateCommand, True)
                ElseIf RowsDeleted = True Then
                    nReturnValue = GetReturnValue(.DeleteCommand)
                    sMessage = GetSprocMessage(.DeleteCommand, True)
                End If
            End If
        End With
        If nReturnValue = 1 Then
            Throw New CPUException(sMessage)
        End If


    End Sub

    Private Function IsDataRowStateChanged(AdapterObj As SqlDataAdapter, TableObj As DataTable, RowstateValue As DataViewRowState) As Boolean
        Dim bChanged As Boolean = False
        Dim bCommandExists As Boolean = False
        With AdapterObj
            Select Case RowstateValue
                Case DataViewRowState.Added
                    If .InsertCommand Is Nothing = False Then
                        bCommandExists = True
                    End If
                Case DataViewRowState.Deleted
                    If .DeleteCommand Is Nothing = False Then
                        bCommandExists = True
                    End If
                Case DataViewRowState.ModifiedCurrent
                    If .UpdateCommand Is Nothing = False Then
                        bCommandExists = True
                    End If
            End Select
        End With

        If bCommandExists = True AndAlso TableObj.Select("", "", RowstateValue).Count > 0 Then
            bChanged = True
        End If

        Return bChanged
    End Function

    Private Sub WriteDebugSQL(AdapterObj As SqlDataAdapter, TableObj As DataTable, ReturnTable As Boolean, RowsAdded As Boolean, RowsModified As Boolean, RowsDeleted As Boolean)
        With AdapterObj
            If ReturnTable = True Then
                If .SelectCommand Is Nothing = False Then
                    Debug.Print(GetSQL(.SelectCommand))
                End If
            Else
                If RowsModified = True Then
                    Debug.Print(GetSQL(.UpdateCommand))
                End If

                If RowsAdded = True Then
                    Debug.Print(GetSQL(.InsertCommand))
                End If

                If RowsDeleted = True Then
                    Debug.Print(GetSQL(.DeleteCommand))
                End If

            End If
        End With

    End Sub

    Private Sub SetAdapterConnection(AdapterObj As SqlDataAdapter, ConnObj As SqlConnection)
        With AdapterObj
            If .SelectCommand Is Nothing = False Then
                .SelectCommand.Connection = ConnObj
            End If
            If .InsertCommand Is Nothing = False Then
                .InsertCommand.Connection = ConnObj
            End If
            If .UpdateCommand Is Nothing = False Then
                .UpdateCommand.Connection = ConnObj
            End If
            If .DeleteCommand Is Nothing = False Then
                .DeleteCommand.Connection = ConnObj
            End If
        End With

    End Sub

    Private Function GetReturnValue(CommandObj As SqlCommand) As Integer
        Dim nReturn As Integer = 0

        For Each objP As SqlParameter In CommandObj.Parameters
            With objP
                If .Direction = ParameterDirection.ReturnValue Then
                    nReturn = .Value
                    Exit For
                End If
            End With
        Next

        Return nReturn
    End Function

    Private Function GetSprocMessage(CommandObj As SqlCommand, SupplyErrorMsgForBlank As Boolean) As String
        Dim sParamName As String = "@vchMessage"
        Dim sMessage As String = ""
        With CommandObj.Parameters
            If .Contains(sParamName) = True AndAlso IsDBNull(.Item(sParamName).Value) = False Then
                sMessage = .Item(sParamName).Value
            End If
        End With
        If sMessage = "" And SupplyErrorMsgForBlank = True Then
            sMessage = "Error Calling " & CommandObj.CommandText
        End If
        Return sMessage
    End Function

    Private Function ParseConstraintViolation(ErrorMessage As String, Prefix As String) As String
        Dim sParsedMessage As String = ErrorMessage
        Dim nPos As Integer = ErrorMessage.IndexOf(Prefix)
        If nPos > -1 Then
            sParsedMessage = sParsedMessage.Substring(nPos + Prefix.Length)
            sParsedMessage = sParsedMessage.Replace("""", "'")
            nPos = sParsedMessage.IndexOf("'")
            If nPos > -1 Then
                sParsedMessage = sParsedMessage.Substring(0, nPos)
            End If
            sParsedMessage = sParsedMessage.Replace("_", " ")
            If Prefix.ToLower = "f_" Then
                nPos = sParsedMessage.IndexOf(" ")
                If nPos > -1 Then
                    sParsedMessage = sParsedMessage.Substring(0, nPos) & " has related records in the " & sParsedMessage.Substring(nPos) & " table"
                End If
                sParsedMessage = "Cannot delete record because " & sParsedMessage
            End If

            sParsedMessage = sParsedMessage.Substring(0, 1).ToUpper & sParsedMessage.Substring(1)

            If sParsedMessage.EndsWith(".") = False Then
                sParsedMessage = sParsedMessage & "."
            End If

        End If


        Return sParsedMessage
    End Function

    Public Function GetFriendlyName(ColumnName As String) As String
        Dim sFriendlyName As String = ColumnName
        If sFriendlyName.ToLower.StartsWith("vch") = True Then
            sFriendlyName = sFriendlyName.Substring(3)
        ElseIf sFriendlyName.ToLower.StartsWith("i") = True Or sFriendlyName.ToLower.StartsWith("m") = True Then
            sFriendlyName = sFriendlyName.Substring(1)
        ElseIf sFriendlyName.ToLower.StartsWith("dt") = True Or sFriendlyName.ToLower.StartsWith("dc") = True Then
            sFriendlyName = sFriendlyName.Substring(2)

        End If
        Dim sWord As String = ""
        For Each sletter As String In sFriendlyName
            If sletter.ToLower <> sletter And sWord > "" Then
                sletter = " " & sletter
            End If
            sWord = sWord & sletter
        Next
        sFriendlyName = Replace(sWord, "  ", " ")
        Return sFriendlyName
    End Function

    Public Function GetResultsToString(TableObj As DataTable, UseFriendlyNames As Boolean) As String
        Dim sValue As String = ""
        Dim sColumns As String = ""
        For Each objc As DataColumn In TableObj.Columns
            If sColumns > "" Then sColumns = sColumns & ","
            If UseFriendlyNames = False Then
                sColumns = sColumns & objc.ColumnName
            Else
                sColumns = sColumns & mdUtility.GetFriendlyName(objc.ColumnName)
            End If
        Next
        sValue = sColumns

        For Each objR As DataRow In TableObj.Rows
            Dim sRowValues As String = ""
            For Each objc As DataColumn In TableObj.Columns
                If sRowValues > "" Then sRowValues = sRowValues & ","
                sRowValues = sRowValues & objR.Item(objc)
            Next
            sValue = sValue & vbCrLf & sRowValues
        Next
        Return sValue
    End Function
    Private Function GetSQL(CommandObj As SqlCommand) As String
        Dim sSQL As String = ""
        Dim sBody As String = ""
        Dim sDeclare As String = "declare @iResult int"
        Dim sSelect As String = "select iResult = @iResult"
        Dim sHeader As String = ""
        Dim sParams As String = ""
        Dim sSetOutputParams = ""
        With CommandObj
            If .Connection Is Nothing = False Then
                sHeader = sHeader & "--" & .Connection.ConnectionString & vbNewLine & vbNewLine
                sHeader = sHeader & "use " & .Connection.Database & vbNewLine & "go" & vbNewLine
            End If

            sBody = "exec @iResult = " & CommandObj.CommandText & vbNewLine

            For Each objP As SqlParameter In .Parameters
                With objP
                    If .Direction = ParameterDirection.InputOutput And IsDBNull(.Value) = False Then
                        sSetOutputParams = sSetOutputParams & .ParameterName & "=" & .Value & ","

                    End If
                End With
            Next

            If sSetOutputParams > "" Then
                sSetOutputParams = "select " + sSetOutputParams
            End If
            For Each objP As SqlParameter In .Parameters
                With objP
                    If .Direction <> ParameterDirection.ReturnValue Then
                        If sParams > "" Then
                            sParams = sParams & "," & vbNewLine
                        End If
                        sParams = sParams & .ParameterName & " = " & GetParamValueForSQL(objP)
                        Select Case .Direction
                            Case ParameterDirection.Output, ParameterDirection.InputOutput
                                sDeclare = sDeclare & ", " & .ParameterName & " " & .SqlDbType.ToString
                                Dim sDimension As String = ""
                                Select Case .SqlDbType
                                    Case SqlDbType.VarChar, SqlDbType.Char
                                        sDimension = " (" & .Size & ")"
                                    Case SqlDbType.Decimal
                                        sDimension = " (" & .Precision & "," & .Scale & ")"
                                End Select
                                sDeclare = sDeclare & sDimension

                                sSelect = sSelect & ", " & .ParameterName.Replace("@", "") & "=" & .ParameterName
                        End Select
                    End If
                End With
            Next

            sBody = sBody & sParams

            sSQL = sHeader & sDeclare & vbNewLine & sSetOutputParams & vbNewLine & sBody & vbNewLine & sSelect
        End With
        Return sSQL
    End Function

    Private Function GetParamValueForSQL(ParamObj As SqlParameter) As String
        Dim sValue As String = ""
        With ParamObj
            If .Direction = ParameterDirection.Output Or .Direction = ParameterDirection.InputOutput Then
                sValue = .ParameterName & " output"
            ElseIf .Value Is Nothing = True OrElse IsDBNull(.Value) = True Then
                sValue = "null"
            Else
                Select Case .SqlDbType
                    Case SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Date, SqlDbType.DateTime
                        sValue = "'" & .Value & "'"
                    Case Else
                        sValue = .Value
                End Select
            End If
        End With
        Return sValue
    End Function
End Module
