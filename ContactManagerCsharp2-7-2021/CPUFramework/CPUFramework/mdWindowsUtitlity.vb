Imports System.Windows.Forms
Public Module mdWindowsUtitlity
    Public Sub FormatReadOnlyGrid(gridobj As DataGridView)
        With gridobj
            .ReadOnly = True
            .AllowUserToAddRows = False
            .RowHeadersWidth = 25
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End With
    End Sub

    Public Sub FormatGridWithFriendlyColumnNames(GridObj As DataGridView)
        For Each objc As DataGridViewColumn In GridObj.Columns
            With objc
                If .HeaderText.ToLower.StartsWith("i") And .HeaderText.ToLower.EndsWith("id") Then
                    .Visible = False
                Else
                    .HeaderText = mdUtility.GetFriendlyName(.HeaderText)
                End If

            End With
        Next
    End Sub

    Public Function GetPrimaryKeyValueFromGrid(GridObj As DataGridView, PrimaryKeyName As String, RowIndex As Integer) As Long
        Dim nId As Long = 0
        If RowIndex > -1 Then
            If GridObj.Columns.Contains(PrimaryKeyName) = False Then
                Throw New Exception("Grid " & GridObj.Name & " does not contain column " & PrimaryKeyName & ". Check your sproc to ensure that it returns primary key name.")
            End If
            nId = GridObj.Item(PrimaryKeyName, RowIndex).Value
        End If
        Return nId
    End Function

    Public Sub GetOpenForm(ByRef FormObj As Form, Optional PrimaryKeyValue As Long = 0)
        For Each objOpenForm As Form In Application.OpenForms
            If objOpenForm.Name = FormObj.Name Then
                If PrimaryKeyValue > 0 Then
                    If objOpenForm.Tag Is Nothing = False AndAlso objOpenForm.Tag = PrimaryKeyValue Then
                        FormObj = objOpenForm
                        Exit For
                    End If
                Else
                    FormObj = objOpenForm
                    Exit For
                End If


            End If
        Next
    End Sub
End Module
