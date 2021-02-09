Imports System.IO
Imports System.Xml
Imports System.Globalization
Imports System.Windows.Forms
Imports Microsoft.Reporting.WinForms
Public Class RdlGenerator
    Public Enum ReportTypeEnum
        None
        Preview
        Excel
    End Enum

    Private objFields As Collection = New Collection
    Dim nColHeaderHeight As Double = 0.17
    Dim nGroupHeaderHeight As Double = 0.32
    Dim nGrandTotalFooterHeight As Double = 0.32
    Dim sFontFamily As String = "Verdana"
    Dim sFontSize As String = "8pt"
    Dim sHeaderFontSize As String = "10pt"
    Public BindingString As String = ""
    Public PageOrientation As PageOrientationEnum = PageOrientationEnum.NotSet
    Public ReportHeader As String = ""
    Dim nPageWidth As Double = 8.5
    Dim nPageHeight As Double = 11
    Dim nSideMargin As Double = 0.5
    Dim nTopMargin As Double = 0.5
    Dim nBottomMargin As Double = 0.5
    Public FormatGLAccountSegments As Boolean = True
    Public RowHeight As Double = 0.15
    Dim bShowBorders As Boolean = False
    Dim bToFile As Boolean = False
    Dim sDataToString As String = ""
    Dim sExtraHeader As String = ""
    Dim objFormat As NumberFormatInfo = New NumberFormatInfo
    Dim nReportType As ReportTypeEnum = ReportTypeEnum.None
    Private Const PivotFieldDelimiter As String = "xqx"
    Public Enum PageOrientationEnum
        NotSet
        Landscape
        Portrait
    End Enum
    Private Enum LinePositionEnum
        Top
        Bottom
    End Enum
    Private Enum TableSectionEnum
        Fields
        TablixCells
        TableCellValues
        Header
        Details
        Footer
        TablixColumns
        TablixColumnHierarchy
        TablixRowHierarchy
        GroupFooter
        GroupHeader
        FooterLiteralText
    End Enum
    Private Enum ValueTypeEnum
        None
        Header
        Value
        Sum
        LiteralText
    End Enum
    Public Class bcGrandTotal
        Public Caption As String
        Public Datafield As String
        Public Total As String
    End Class
    'Public Function Run(ByVal ReportHeaderValue As String, ByVal BCGridObject As bcGrid, ByVal ReportType As ReportTypeEnum, Optional ByVal ExcludeList As Collection = Nothing, Optional ByVal IncludeCriteria As Boolean = False) As MemoryStream
    '    Return Me.GenerateRDLFromGrid(ReportHeaderValue, BCGridObject, ReportType, ExcludeList, IncludeCriteria)
    'End Function
    Public Function Run(ByVal ReportHeaderValue As String, ByVal TableObject As DataTable, ByVal BindingStringValue As String, ByVal ReportType As ReportTypeEnum, Optional ByVal ExcludeList As Collection = Nothing, Optional ByVal Criteria As String = "", Optional ByVal GroupField As String = "", Optional ByVal ExtraHeader As String = "", Optional ByVal TotalCols As String = "", Optional ByVal GrandTotalCols As String = "") As MemoryStream
        If BindingStringValue <> "" Then Me.BindingString = BindingStringValue
        Return Me.GenerateRDLFromTable(ReportHeaderValue, TableObject, ReportType, ExcludeList, Criteria, GroupField, ExtraHeader, TotalCols, GrandTotalCols)
    End Function

    Private Sub DoAddField(ByVal Field As String, ByVal HeaderText As String, ByVal WidthValue As Decimal, ByVal Position As Integer, ByVal UseCurrencyFormat As Boolean, ByVal FieldDataType As Type, ByVal TextAlignment As DataGridViewContentAlignment, ByVal ExcludeList As Collection, ByVal SumColumnsCol As Collection, ByVal OnlyIfExists As Boolean, ByVal TotalColumnsCol As Collection, ByVal GrandTotalColumnsCol As Collection, Optional CustomWidth As Decimal = -1)
        Dim sField As String = Field
        Dim objSumColumnsCol As Collection = SumColumnsCol
        If ExcludeList Is Nothing = False And HeaderText <> "" Then
            If ExcludeList.Contains(HeaderText) = True Then
                Exit Sub
            End If
        ElseIf ExcludeList Is Nothing = False Then
            If ExcludeList.Contains(Field) = True Then
                Exit Sub
            End If
        End If
        Dim objField As bcField = Nothing
        If objFields.Contains(sField) = True Then
            objField = objFields(sField)
            If objField.Visible = False Then
                Dim sKey As String = objField.Key
                objField = Nothing
                objFields.Remove(sKey)
                Exit Sub
            End If
        Else
            If OnlyIfExists = True Then Exit Sub
            objField = New bcField
        End If
        'Position = Position + 1
        With objField
            .Key = sField
            If .Header = "" Then .Header = Replace(HeaderText, vbCrLf, " ")
            .Width = WidthValue
            If CustomWidth > -1 Then
                .CustomWidth = CustomWidth
            End If
            If .CustomPosition = 0 Then
                .Position = Position
            Else
                .Position = .CustomPosition
            End If
        End With
        If UseCurrencyFormat = True Then
            objField.IsCurrency = True
        Else
            If FieldDataType Is GetType(Date) = True Then
                objField.IsDate = True
            End If
        End If
        If objSumColumnsCol Is Nothing = False Then
            If objSumColumnsCol.Contains(objField.Key) Then
                Dim bSum As Boolean = True
                Dim bGrand As Boolean = True
                If TotalColumnsCol Is Nothing = False Then
                    bSum = Me.GetSegmentValue(objField.Key, TotalColumnsCol)
                End If
                If GrandTotalColumnsCol Is Nothing = False Then
                    bGrand = Me.GetSegmentValue(objField.Key, GrandTotalColumnsCol)
                End If
                If bSum = True Then objField.Sum = True
                If bGrand = True Then objField.GrantTotalSum = True
            End If
        End If
        Select Case TextAlignment
            Case DataGridViewContentAlignment.MiddleRight, DataGridViewContentAlignment.BottomRight, DataGridViewContentAlignment.TopRight
                objField.TextAlignment = "Right"
            Case Else
                objField.TextAlignment = "Left"
        End Select
        If objFields.Contains(objField.Key) = False Then
            objFields.Add(objField, objField.Key)
        End If
    End Sub

    'Private Function GenerateRDLFromGrid(ByVal ReportHeaderValue As String, ByVal BCGridObject As bcGrid, ByVal ReportType As ReportTypeEnum, Optional ByVal ExcludeList As Collection = Nothing, Optional ByVal IncludeCriteria As Boolean = False, Optional ByVal GroupField As String = "") As MemoryStream
    '    nReportType = ReportType
    '    Dim objMemoryStream As IO.MemoryStream
    '    If ReportHeader = "" Then ReportHeader = ReportHeaderValue
    '    Dim objGrid As DataGridView = BCGridObject.Grid
    '    Dim objTotalCol As Collection = BCGridObject.GrandTotals(False)
    '    Dim nPosition As Integer = 0
    '    Dim objTable As DataTable = Nothing
    '    Dim sCriteria As String = ""
    '    Dim bOnlyIfExists As Boolean = False
    '    If IncludeCriteria = True Then
    '        sCriteria = BCGridObject.Criteria
    '    End If
    '    If TypeOf (objGrid.DataSource) Is DataTable Then
    '        objTable = objGrid.DataSource
    '    Else
    '        Throw New Exception("Report is only implemented for grid with datatable as source")
    '    End If
    '    If BindingString.EndsWith("#") Then
    '        BindingString = BindingString.Substring(0, BindingString.Length - 1)
    '        bOnlyIfExists = True
    '    End If
    '    If BindingString <> "" Then Me.AddFieldsFromBindingString(BindingString, ExcludeList)
    '    For Each objC As DataGridViewColumn In objGrid.Columns
    '        If objC.Visible = True Then
    '            nPosition = nPosition + 1
    '            Dim sField As String = objC.DataPropertyName
    '            Dim objR As Rectangle = objGrid.GetColumnDisplayRectangle(objC.Index, False)
    '            Dim nWidth As Decimal = objR.Width / 75
    '            Dim nCustomWidth As Decimal = -1
    '            If objC.Tag Is Nothing = False AndAlso IsNumeric(objC.Tag.ToString) = True Then
    '                nWidth = objC.Tag.ToString
    '                nCustomWidth = nWidth
    '                'nWidth = nWidth 
    '            End If
    '            Dim bCurrency As Boolean = objC.DefaultCellStyle.Format = Me.CurrencyFormat(True)
    '            Dim nTextAlignment As DataGridViewContentAlignment = objC.DefaultCellStyle.Alignment
    '            Dim nType As Type = Nothing
    '            With objTable
    '                If .Columns.Contains(sField) = True Then
    '                    nType = .Columns(sField).DataType
    '                Else
    '                    nType = GetType(String)
    '                End If
    '            End With
    '            Me.DoAddField(sField, objC.HeaderText, nWidth, nPosition, bCurrency, nType, nTextAlignment, ExcludeList, objTotalCol, bOnlyIfExists = True, TotalColumnsCol:=Nothing, GrandTotalColumnsCol:=Nothing, CustomWidth:=nCustomWidth)
    '        End If
    '    Next
    '    objMemoryStream = Me.DoGenerateRdl(ReportType, sCriteria, GroupField)
    '    sDataToString = Me.GetDataToString(objTable)
    '    Return objMemoryStream
    'End Function

    Private Function GenerateRDLFromTable(ByVal ReportHeaderValue As String, ByVal TableObject As DataTable, ByVal ReportType As ReportTypeEnum, Optional ByVal ExcludeList As Collection = Nothing, Optional ByVal Criteria As String = "", Optional ByVal GroupField As String = "", Optional ByVal ExtraHeaderValue As String = "", Optional ByVal TotalCols As String = "", Optional ByVal GrandTotalCols As String = "") As MemoryStream
        nReportType = ReportType
        Dim objMemoryStream As IO.MemoryStream
        Dim bOnlyIfExists As Boolean = False
        If ReportHeader = "" Then ReportHeader = ReportHeaderValue
        sExtraHeader = ExtraHeaderValue
        Dim objSumColumnsCol As Collection = Me.GetGrandTotalsFromTable(TableObject)
        Dim nPosition As Integer = 0
        Dim objTable As DataTable = Nothing
        objTable = TableObject
        ' nColHeaderHeight = 0.4 'objGrid.ColumnHeadersHeight / 68
        'If BindingString.EndsWith("#") Then
        If BindingString > "" Then
            BindingString = BindingString.Substring(0, BindingString.Length - 1)
            bOnlyIfExists = True
        End If
        If BindingString <> "" Then Me.AddFieldsFromBindingString(BindingString, ExcludeList)
        Dim objTotalColumnsCol As Collection = New Collection
        For Each sCol As String In TotalCols.Split(",")
            If sCol.Trim <> "" Then
                Dim sKey() As String = sCol.Split(":")
                If objTotalColumnsCol.Contains(sKey(0)) = False Then
                    objTotalColumnsCol.Add(sKey(0), sCol)
                End If
            End If
        Next
        Dim objGrandTotalColumnsCol As Collection = New Collection
        For Each sCol As String In GrandTotalCols.Split(",")
            If sCol.Trim <> "" Then
                Dim sKey() As String = sCol.Split(":")
                If objGrandTotalColumnsCol.Contains(sKey(0)) = False Then
                    objGrandTotalColumnsCol.Add(sKey(0), sCol)
                End If
            End If
        Next
        For Each objF As DataColumn In objTable.Columns
            nPosition = nPosition + 1
            Dim sField As String = objF.ColumnName
            Dim sHeader As String = Me.GetEnglishNameForField(sField)
            'Dim objR As Rectangle = objGrid.GetColumnDisplayRectangle(objC.Index, False)
            Dim nWidth As Decimal = 0 'objR.Width / 75
            Dim bCurrency As Boolean = False
            Dim nTextAlignment As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft
            Dim nType As Type = Nothing
            With objTable
                If .Columns.Contains(sField) = True Then
                    nType = .Columns(sField).DataType
                Else
                    nType = GetType(String)
                End If
            End With
            If nType Is GetType(Decimal) Or nType Is GetType(Double) Then
                bCurrency = True
                nTextAlignment = DataGridViewContentAlignment.MiddleRight
                nWidth = 1.5
            ElseIf nType Is GetType(String) Then
                nWidth = 1.5
            Else
                nWidth = 0.85
            End If
            Me.DoAddField(sField, sHeader, nWidth, nPosition, bCurrency, nType, nTextAlignment, ExcludeList, objSumColumnsCol, bOnlyIfExists, objTotalColumnsCol, objGrandTotalColumnsCol)
        Next
        sDataToString = Me.GetDataToString(TableObject)
        If Criteria <> "" Then Criteria = vbCrLf & Criteria
        objMemoryStream = Me.DoGenerateRdl(ReportType, Criteria, GroupField)
        Return objMemoryStream
    End Function
    Private Function DoGenerateRdl(ByVal ReportType As ReportTypeEnum, Optional ByVal LiteralText As String = "", Optional ByVal GroupField As String = "") As MemoryStream
        nReportType = ReportType
        'If nReportType = ReportTypeEnum.Excel Then
        '    LiteralText = "" 'causes merged columns which messes up filtering
        '    GroupField = ""
        'End If
        'Dim nReportWidth As Integer = 6
        Dim objMemoryStream As IO.MemoryStream = New MemoryStream()
        Dim objWriter As XmlWriter
        Dim sGroupFields() As String = GroupField.Split(":")
        Dim sMainGroup As String = ""
        Dim sFile As String = ""
        Dim objSettings As XmlWriterSettings = New XmlWriterSettings()
        Dim nFooterLeft As Double = 0
        If sGroupFields.Length > 0 Then
            sMainGroup = sGroupFields(0)
        End If
        If Me.FormatGLAccountSegments = True Then Me.DoFormatGLAccountSegments()
        Me.SetFieldWidths()
        Me.SetPageOrientation()
        If PageOrientation = PageOrientationEnum.Landscape Then
            Dim n As Double = nPageWidth
            nPageWidth = nPageHeight
            nPageHeight = n
        End If
        objSettings.Indent = True
        objSettings.OmitXmlDeclaration = True
        objSettings.NewLineOnAttributes = True

        objWriter = XmlWriter.Create(objMemoryStream, objSettings)
        If bToFile = True Then
            '           objWriterToFile = XmlWriter.Create("\\wcrsana2\programmers$\mgellis\test" & Now.Second & ".xml", objSettings)
            sFile = "\\wcrsana2\programmers$\mgellis\test" & Now.Second & ".xml"
        End If

        With objWriter
            ' Report element
            .WriteProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
            '<Report xmlns:rd="http://schemas.microsoft.com/SQLServer/reporting/reportdesigner" xmlns="http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition">
            .WriteStartElement("Report", "http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition")
            '.WriteAttributeString("rd", "http://schemas.microsoft.com/SQLServer/reporting/reportdesigner")
            .WriteElementString("Width", Me.ReportWidth.ToString(objFormat) & "in")
            .WriteElementString("Language", "=User!Language")
            ' DataSource element
            .WriteStartElement("DataSources")
            .WriteStartElement("DataSource")
            .WriteAttributeString("Name", Nothing, "DataSource1")
            .WriteStartElement("ConnectionProperties")
            .WriteElementString("DataProvider", "SQL")
            .WriteElementString("ConnectString", Nothing)
            .WriteElementString("IntegratedSecurity", "true")
            .WriteEndElement() ' ConnectionProperties
            .WriteEndElement() ' DataSource
            .WriteEndElement() ' DataSources
            ' DataSet element
            .WriteStartElement("DataSets")
            .WriteStartElement("DataSet")
            .WriteAttributeString("Name", Nothing, "DataSet1")
            ' Query element
            .WriteStartElement("Query")
            .WriteElementString("DataSourceName", "DataSource1")
            .WriteElementString("CommandText", Nothing)
            .WriteElementString("UseGenericDesigner", "rd", True)
            .WriteEndElement() ' Query

            Me.AddTableSection(TableSectionEnum.Fields, objWriter)

            .WriteEndElement() ' DataSet
            .WriteEndElement() ' DataSets
            .WriteStartElement("Page")
            .WriteElementString("PageHeight", nPageHeight.ToString(objFormat) & "in")
            .WriteElementString("PageWidth", nPageWidth.ToString(objFormat) & "in")
            .WriteElementString("RightMargin", nSideMargin.ToString(objFormat) & "in")
            .WriteElementString("LeftMargin", nSideMargin.ToString(objFormat) & "in")
            .WriteElementString("TopMargin", nTopMargin.ToString(objFormat) & "in")
            .WriteElementString("BottomMargin", nBottomMargin.ToString(objFormat) & "in")
            'header
            'If nReportType <> ReportTypeEnum.Excel Then
            .WriteStartElement("PageHeader")
            .WriteElementString("PrintOnFirstPage", "true")
            .WriteStartElement("Style")
            Me.AddSectionLine(objWriter, LinePositionEnum.Bottom)
            .WriteEndElement() '      </Style>


            .WriteStartElement("ReportItems")
            If sExtraHeader <> "" Then
                .WriteStartElement("Textbox")
                .WriteAttributeString("Name", "txtExtraHeader")
                .WriteElementString("Width", (Me.ReportWidth - 1).ToString(objFormat) & "in")
                .WriteElementString("Top", ".175in")
                .WriteElementString("CanGrow", "true")
                .WriteElementString("CanShrink", "false")
                .WriteElementString("Left", "0in")
                Me.AddTextboxValue(objWriter, New bcTextbox("ReplaceExtraHeader", sFontFamily, sHeaderFontSize, "", "", "", ""), False)
                .WriteEndElement() '</Textbox>
            End If

            .WriteStartElement("Textbox")
            .WriteAttributeString("Name", "txtReportHeader")
            .WriteElementString("Width", (Me.ReportWidth - 1).ToString(objFormat) & "in")
            .WriteElementString("CanGrow", "true")
            .WriteElementString("CanShrink", "false")
            .WriteElementString("Left", "0in")
            Me.AddTextboxValue(objWriter, New bcTextbox(ReportHeader, sFontFamily, sHeaderFontSize, "", "", "", ""), False)
            .WriteEndElement() '</Textbox>
            .WriteEndElement() '</ReportItems>
            If sExtraHeader <> "" Then
                Dim sHeight As String = "0.50in"
                Dim nHeight As Decimal = Me.GetLineCount(sExtraHeader)
                If nHeight > 2 Then
                    nHeight = Math.Round(nHeight / 4.5, 2)
                    sHeight = nHeight.ToString & "in"
                End If
                .WriteElementString("Height", sHeight)
            Else
                .WriteElementString("Height", "0.25in")
            End If

            .WriteElementString("PrintOnLastPage", "true")
            .WriteEndElement() '</PageHeader>
            'End If
            'footer
            .WriteStartElement("PageFooter")
            .WriteElementString("PrintOnFirstPage", "true")
            .WriteStartElement("Style")
            Me.AddSectionPadding(objWriter)
            Me.AddSectionLine(objWriter, LinePositionEnum.Top)
            .WriteEndElement() '      </Style>
            'If ReportType <> ReportTypeEnum.Excel Then
            .WriteStartElement("ReportItems")
            'page number
            .WriteStartElement("Textbox")
            .WriteAttributeString("Name", "txtFooter")
            .WriteElementString("Width", (Me.ReportWidth - 1).ToString(objFormat) & "in")
            .WriteElementString("Top", ".08in")
            .WriteElementString("CanGrow", "true")
            .WriteElementString("CanShrink", "false")
            .WriteElementString("Left", nFooterLeft.ToString(objFormat) & "in")
            Dim s As String = "=" & Chr(34) & "Page " & Chr(34) & " &  Globals!PageNumber & " & Chr(34) & " of " & Chr(34) & " & Globals!TotalPages &"
            s = s & Chr(34) & ", Printed on  " & Chr(34) & " & FormatDateTime(Globals!ExecutionTime,2)"
            Me.AddTextboxValue(objWriter, New bcTextbox(s, sFontFamily, sFontSize, "", "", "", "Left"), False)
            .WriteEndElement() '    </Textbox>
            .WriteEndElement() '  </ReportItems>
            'End If
            .WriteElementString("Height", ".35in")
            .WriteElementString("PrintOnLastPage", "true")

            .WriteEndElement() '</PageFooter>
            .WriteEndElement() 'Page
            .WriteElementString("ConsumeContainerWhitespace", "true")
            .WriteStartElement("Body")
            .WriteElementString("Height", Me.ReportHeight.ToString(objFormat) & "in")

            ' ReportItems element
            .WriteStartElement("ReportItems")

            ' Table element
            .WriteStartElement("Tablix")
            .WriteAttributeString("Name", "Table1")
            .WriteStartElement("TablixBody")
            'columns
            Me.AddTableSection(TableSectionEnum.TablixColumns, objWriter)
            'rows
            .WriteStartElement("TablixRows")
            'header row
            .WriteStartElement("TablixRow")
            .WriteElementString("Height", nColHeaderHeight.ToString(objFormat) & "in")
            Me.AddTableSection(TableSectionEnum.TablixCells, objWriter)
            .WriteEndElement() ' TableRow
            If GroupField <> "" Then
                'group row
                .WriteStartElement("TablixRow")
                .WriteElementString("Height", nGroupHeaderHeight.ToString(objFormat) & "in")
                Me.AddTableSection(TableSectionEnum.GroupHeader, objWriter, GroupName:="Group" & sMainGroup, GroupField:=sMainGroup)
                .WriteEndElement() ' TableRow
                If sGroupFields.Length > 1 Then
                    Dim sSubGroup As String = sGroupFields(1)
                    .WriteStartElement("TablixRow")
                    .WriteElementString("Height", Me.RowHeight.ToString(objFormat) & "in")
                    Me.AddTableSection(TableSectionEnum.GroupHeader, objWriter, GroupName:="Group" & sSubGroup, GroupField:=sSubGroup, SubGroup:=True)
                    .WriteEndElement() ' TableRow
                End If
            End If
            'value row
            .WriteStartElement("TablixRow")
            .WriteElementString("Height", Me.RowHeight.ToString(objFormat) & "in")
            Me.AddTableSection(TableSectionEnum.TableCellValues, objWriter)
            .WriteEndElement() ' TableRow
            If GroupField <> "" Then
                'footer group row
                If sGroupFields.Length > 1 Then
                    Dim sSubGroup As String = sGroupFields(1)
                    .WriteStartElement("TablixRow")
                    .WriteElementString("Height", (Me.RowHeight * 1.5).ToString(objFormat) & "in")
                    Me.AddTableSection(TableSectionEnum.GroupFooter, objWriter, GroupName:="Group" & sSubGroup, GroupField:=sSubGroup, SubGroup:=True)
                    .WriteEndElement() ' TableRow
                End If

                .WriteStartElement("TablixRow")
                .WriteElementString("Height", nGroupHeaderHeight.ToString(objFormat) & "in")
                Me.AddTableSection(TableSectionEnum.GroupFooter, objWriter, GroupName:="Group" & sMainGroup, GroupField:=sMainGroup)
                .WriteEndElement() ' TableRow
            End If
            'footer grand total
            Dim bHasGrouping As Boolean = False
            If GroupField <> "" Then bHasGrouping = True
            .WriteStartElement("TablixRow")
            .WriteElementString("Height", nGrandTotalFooterHeight.ToString(objFormat) & "in")
            Me.AddTableSection(TableSectionEnum.Footer, objWriter, ContainsGrouping:=bHasGrouping)
            .WriteEndElement() ' TableRow

            'literal text
            If LiteralText <> "" Then
                .WriteStartElement("TablixRow")
                .WriteElementString("Height", nGroupHeaderHeight.ToString(objFormat) & "in")
                Me.AddTableSection(TableSectionEnum.FooterLiteralText, objWriter, LiteralText)
                .WriteEndElement() ' TableRow
            End If
            .WriteEndElement() ' TableRows
            .WriteEndElement() 'TablixBody

            Me.AddTableSection(TableSectionEnum.TablixColumnHierarchy, objWriter)

            .WriteStartElement("TablixRowHierarchy")
            .WriteStartElement("TablixMembers")
            .WriteStartElement("TablixMember") 'header row
            .WriteElementString("RepeatOnNewPage", "true")
            .WriteElementString("KeepTogether", "true")
            .WriteElementString("KeepWithGroup", "After")
            .WriteEndElement() 'TablixMember
            'group by 
            If GroupField <> "" Then
                If sGroupFields.Length < 2 Then
                    .WriteStartElement("TablixMember") 'main
                    .WriteStartElement("Group")
                    .WriteAttributeString("Name", "Group" & sMainGroup)
                    .WriteStartElement("GroupExpressions")
                    .WriteElementString("GroupExpression", "=Fields!" & sMainGroup & ".Value")
                    .WriteEndElement() 'GroupExpressions
                    .WriteEndElement() 'Group

                    .WriteStartElement("TablixMembers")
                    .WriteStartElement("TablixMember")
                    .WriteElementString("KeepWithGroup", "After")
                    .WriteElementString("RepeatOnNewPage", "true")
                    .WriteEndElement()
                    .WriteStartElement("TablixMember")
                    .WriteStartElement("Group")
                    .WriteAttributeString("Name", "DetailGroupBy")
                    .WriteEndElement() 'Group
                    .WriteElementString("KeepTogether", "true")
                    .WriteEndElement() 'TablixMember

                    .WriteStartElement("TablixMember") 'for group footer
                    .WriteElementString("KeepWithGroup", "Before")
                    .WriteElementString("KeepTogether", "true")
                    .WriteEndElement() '

                    .WriteEndElement() 'TablixMembers
                    .WriteEndElement() 'TabblixMember main
                Else
                    Dim sSubGroup As String = sGroupFields(1)
                    .WriteStartElement("TablixMember") 'main
                    .WriteStartElement("Group")
                    .WriteAttributeString("Name", "Group" & sMainGroup)
                    .WriteStartElement("GroupExpressions")
                    .WriteElementString("GroupExpression", "=Fields!" & sMainGroup & ".Value")
                    .WriteEndElement() 'GroupExpressions
                    .WriteEndElement() 'Group

                    .WriteStartElement("TablixMembers")
                    'blank tablix member
                    .WriteStartElement("TablixMember")
                    .WriteEndElement()
                    'sub group
                    .WriteStartElement("TablixMember")
                    .WriteStartElement("Group")
                    .WriteAttributeString("Name", "Group" & sSubGroup)
                    .WriteStartElement("GroupExpressions")
                    .WriteElementString("GroupExpression", "=Fields!" & sSubGroup & ".Value")
                    .WriteEndElement() 'GroupExpressions
                    .WriteEndElement() 'Group

                    .WriteStartElement("TablixMembers")
                    .WriteStartElement("TablixMember")
                    .WriteElementString("KeepWithGroup", "After")
                    .WriteEndElement()  'TablixMember

                    .WriteStartElement("TablixMember")
                    .WriteStartElement("Group")
                    .WriteAttributeString("Name", "DetailGroupBy")
                    .WriteEndElement() 'Group
                    .WriteElementString("KeepTogether", "true")
                    .WriteEndElement() 'TablixMember

                    .WriteStartElement("TablixMember")
                    .WriteElementString("KeepWithGroup", "Before")
                    .WriteEndElement()  'TablixMembers
                    .WriteEndElement()  'TablixMember
                    .WriteEndElement() '
                End If
            Else
                .WriteStartElement("TablixMember") 'value row
                .WriteStartElement("Group")
                .WriteAttributeString("Name", "DetailGroup")
                .WriteEndElement() 'Group
                .WriteEndElement() 'TablixMember
            End If
            'grand total row
            .WriteStartElement("TablixMember")
            Dim sKeep As String = "Before"
            .WriteElementString("KeepWithGroup", sKeep)
            .WriteElementString("KeepTogether", "true")
            .WriteEndElement() 'grand total row

            If LiteralText <> "" Then
                .WriteStartElement("TablixMember") 'literal text
                .WriteElementString("KeepWithGroup", "Before")
                .WriteElementString("KeepTogether", "true")
                .WriteEndElement() 'tablixmember
            End If
            .WriteEndElement() 'TablixMembers
            .WriteEndElement() 'TablixRowHierarchy
            If sGroupFields.Length > 1 Then
                .WriteStartElement("TablixMember")
                .WriteElementString("KeepTogether", "true")
                .WriteEndElement() 'tablixmember
                .WriteEndElement()
                .WriteEndElement()
            End If

            .WriteEndElement() 'Tablix
            .WriteEndElement() ' ReportItems
            .WriteEndElement() ' Body
            .WriteEndElement() ' Report
            ' Flush the writer and close the stream
            .Flush()
            objMemoryStream.Flush()
            objMemoryStream.Position = 0
        End With
        If bToFile = True Then
            Dim o As MemoryStream = New MemoryStream
            objMemoryStream.CopyTo(o)
            objMemoryStream.Position = 0
            If IO.File.Exists(sFile) = True Then
                IO.File.Delete(sFile)
            End If
            Dim objF As FileStream = IO.File.Create(sFile)
            objF.Write(o.ToArray, 0, o.Length)
            objF.Close()
        End If
        Return objMemoryStream
    End Function


    Private Sub AddTableSection(ByVal TableSection As TableSectionEnum, ByVal WriterObject As XmlWriter, Optional ByVal LiteralText As String = "", Optional ByVal GroupName As String = "", Optional ByVal GroupField As String = "", Optional ByVal SubGroup As Boolean = False, Optional ByVal ContainsGrouping As Boolean = False)
        Dim objElementCol As Collection = New Collection
        Dim bWidth As Boolean = False
        Dim bName As Boolean = False
        Dim bDataField As Boolean = False
        Dim bUnderline As Boolean = False
        Dim bBold As Boolean = False
        Dim bStyle As Boolean = False
        Dim bPosition As Boolean = False
        Dim bTextAlign As Boolean = False
        Dim sNamePrefix As String = ""
        Dim bCurrencyFormat As Boolean = False
        Dim sMainElement As String = TableSection.ToString
        Dim nValueType As ValueTypeEnum = ValueTypeEnum.None
        Dim bVerticalBottom As Boolean = False
        Dim bPadding = False
        Dim bMembers As Boolean = False
        Dim nColspan As Integer = 0
        Dim bVisibility As Boolean = False
        Dim bContainsGrandTotal As Boolean = Me.FieldsContainGrandTotal
        Dim bContainsSum As Boolean = Me.FieldsContainSum
        Dim bAlternateBackColor As Boolean
        '  sNamePrefix = TableSection.ToString
        Select Case TableSection
            Case TableSectionEnum.TablixColumns
                objElementCol.Add("TablixColumn")
                bWidth = True
            Case TableSectionEnum.Fields
                objElementCol.Add("Field")
                bName = True
                bDataField = True
            Case TableSectionEnum.TablixCells
                objElementCol.Add("TablixCell")
                objElementCol.Add("CellContents")
                objElementCol.Add("Textbox")
                nValueType = ValueTypeEnum.Header
                bName = True
                bStyle = True
                bUnderline = True
                bBold = True
                bTextAlign = True
                bPosition = True
                sNamePrefix = TableSection.ToString  '"Header"
            Case TableSectionEnum.GroupHeader
                sMainElement = "TablixCells"
                objElementCol.Add("TablixCell")
                objElementCol.Add("CellContents")
                objElementCol.Add("Textbox")
                nValueType = ValueTypeEnum.Header
                bName = True
                bStyle = True
                bBold = True
                bTextAlign = True
                bPosition = True
                sNamePrefix = TableSection.ToString
                bVerticalBottom = True
            Case TableSectionEnum.TableCellValues
                sMainElement = "TablixCells"
                objElementCol.Add("TablixCell")
                objElementCol.Add("CellContents")
                objElementCol.Add("Textbox")
                bName = True
                bStyle = True
                bTextAlign = True
                bCurrencyFormat = True
                bPosition = True
                nValueType = ValueTypeEnum.Value
                bPadding = True
                bAlternateBackColor = True
            Case TableSectionEnum.FooterLiteralText
                sMainElement = "TablixCells"
                objElementCol.Add("TablixCell")
                objElementCol.Add("CellContents")
                objElementCol.Add("Textbox")
                bName = True
                bStyle = True
                bTextAlign = True
                bCurrencyFormat = False
                bPosition = True
                nValueType = ValueTypeEnum.LiteralText
                bWidth = False
                sNamePrefix = "Criteria"
            Case TableSectionEnum.GroupFooter
                sMainElement = "TablixCells"
                objElementCol.Add("TablixCell")
                objElementCol.Add("CellContents")
                objElementCol.Add("Textbox")
                bName = True
                sNamePrefix = TableSection.ToString
                bCurrencyFormat = True
                bStyle = True
                bTextAlign = True
                bPosition = True
                nValueType = ValueTypeEnum.Sum
                If SubGroup = False Then
                    bBold = True
                End If
            Case TableSectionEnum.Footer
                sMainElement = "TablixCells"
                objElementCol.Add("TablixCell")
                objElementCol.Add("CellContents")
                objElementCol.Add("Textbox")
                bName = True
                sNamePrefix = TableSection.ToString
                bCurrencyFormat = True
                bStyle = True
                bTextAlign = True
                bPosition = True
                nValueType = ValueTypeEnum.Sum
                bBold = True
                'bUnderline = True
                bVerticalBottom = True
                ' bShowBorders = True
            Case TableSectionEnum.TablixColumnHierarchy
                sMainElement = "TablixColumnHierarchy"
                objElementCol.Add("TablixMember")
                bMembers = True
                bVisibility = True
            Case TableSectionEnum.TablixRowHierarchy
                sMainElement = "TablixRowHierarchy"
                objElementCol.Add("TablixMember")
                bMembers = True
        End Select
        With WriterObject
            .WriteStartElement(sMainElement)
            If bMembers = True Then
                .WriteStartElement("TablixMembers")
            End If
            Dim objFieldName As bcField = Me.GetNextField(Nothing)
            Dim nPos As Integer = 0
            Do Until objFieldName Is Nothing = True
                nPos = objFieldName.Position
                Dim bTextbox As Boolean = False
                For Each objE As String In objElementCol
                    .WriteStartElement(objE)
                    If objE = "Textbox" Then
                        bTextbox = True
                    Else
                        bTextbox = False
                    End If
                Next objE
                Dim sName As String = sNamePrefix & objFieldName.Key
                If SubGroup = True Then sName = sNamePrefix & GroupField & nPos
                If bName = True Then .WriteAttributeString("Name", Nothing, sName)
                If bWidth = True Or bVisibility = True Then
                    Dim nWidth As Decimal = objFieldName.Width
                    If nWidth < 0 Then nWidth = 0
                    If bWidth = True Then
                        Debug.Print(sName & "  " & nWidth)
                        .WriteElementString("Width", nWidth.ToString(objFormat) & "in")

                    End If
                End If
                If bTextbox = True Then
                    Dim objTxt As bcTextbox = New bcTextbox
                    Dim bCanShrink As Boolean = True
                    With objTxt
                        If TableSection = TableSectionEnum.Footer Or TableSection = TableSectionEnum.TablixCells Then
                            .CanGrow = "true"
                        ElseIf objFieldName.IsCurrency = True Or objFieldName.IsDate = True Or TableSection = TableSectionEnum.FooterLiteralText Then
                            .CanGrow = "true"
                        ElseIf TableSection = TableSectionEnum.GroupFooter And SubGroup = True And nPos = 2 Then
                            .CanGrow = "true"
                            bCanShrink = False
                        Else
                            .CanGrow = ""
                        End If
                        .CanGrow = "true"
                        bCanShrink = False
                        'If bCanShrink = True Then .CanShrink = "true"
                        If TableSection = TableSectionEnum.FooterLiteralText Or TableSection = TableSectionEnum.GroupHeader Then
                            nColspan = objFields.Count 'only one cell (exit loop at bottom) that spans table
                        End If
                        If bStyle = True Then
                            .FontFamily = sFontFamily
                            .FontSize = sFontSize
                            '---box around textbox to troubleshoot spacing
                            If nReportType = ReportTypeEnum.Excel Or (bShowBorders = True And bTextbox = True And TableSection = TableSectionEnum.TableCellValues) Then
                                .BorderAll = True
                                .BorderColor = "Silver"
                                .BorderStyle = "Solid"
                                .BorderWidth = "1pt"
                                '.WriteElementString("Left", "1pt")
                            End If
                            If bTextbox = True And objFieldName.TextAlignment.ToLower = "right" Then
                                '.WriteElementString("PaddingLeft", "2pt")
                                .PaddingRightWidth = "10pt"
                            ElseIf bPadding = True Then
                                .PaddingRightWidth = objFieldName.Padding & "pt"
                            ElseIf SubGroup = True And (TableSection = TableSectionEnum.GroupHeader Or (TableSection = TableSectionEnum.GroupFooter And nPos = 2)) Then
                                .PaddingLeftWidth = "20pt"
                            End If
                            If bUnderline = True Then
                                .TextDecoration = "Underline"
                            End If
                            If bBold = True Then
                                .FontWeight = "Bold"
                            End If
                            If nValueType = ValueTypeEnum.Sum And ((objFieldName.Sum = True And TableSection <> TableSectionEnum.Footer) Or (objFieldName.GrantTotalSum = True And TableSection = TableSectionEnum.Footer)) Then
                                .BorderTop = True
                                .BorderStyle = "Solid"
                                .BorderTopWidth = "1pt"
                                Select Case TableSection
                                    Case TableSectionEnum.Footer
                                        .BorderBottom = True
                                        .BorderBottomWidth = "2pt"
                                    Case TableSectionEnum.GroupFooter
                                        .BorderTopWidth = "1pt"
                                End Select
                            End If
                            If bTextAlign = True Then .TextAlign = objFieldName.TextAlignment
                            If bVerticalBottom = True Then .VerticalAlign = "Bottom"
                            If bCurrencyFormat = True Then
                                If objFieldName.IsCurrency = True Then
                                    Dim sField As String = ""
                                    If nValueType <> ValueTypeEnum.Sum Then
                                        sField = "Fields!" & objFieldName.Key & ".Value"
                                    Else
                                        sField = "Sum(Fields!" & objFieldName.Key & ".Value)"
                                    End If
                                    .Color = "=IIF(" & sField & "<0," & Chr(34) & "Red" & Chr(34) & "," & Chr(34) & "Black" & Chr(34) & ")"
                                End If
                            End If
                            .AlternateBackColor = bAlternateBackColor
                        End If
                        If bPosition = True Then
                            '.WriteElementString("Top", "0in")
                            '.WriteElementString("Left", "0in")
                        End If

                        If nValueType <> ValueTypeEnum.None Or (TableSection = TableSectionEnum.GroupFooter And SubGroup = True And nPos = 2) Then
                            Dim sValue As String = ""
                            If bContainsSum = True And TableSection = TableSectionEnum.GroupFooter And SubGroup = True And nPos = 2 Then
                                sValue = "" '"=" & """Subtotal"""
                            Else
                                Select Case nValueType
                                    Case ValueTypeEnum.Header
                                        If TableSection = TableSectionEnum.GroupHeader Then
                                            If objFieldName.Key.ToLower = GroupField.ToLower Then
                                                sValue = "=Fields!" & GroupField & ".Value"
                                                objTxt.DefaultName = "Group1"
                                            ElseIf SubGroup = True Then
                                                sValue = "=Fields!" & GroupField & ".Value"
                                                objTxt.DefaultName = "Group2"
                                            Else
                                                sValue = ""
                                            End If
                                        Else
                                            sValue = objFieldName.Header
                                        End If
                                    Case ValueTypeEnum.Value, ValueTypeEnum.Sum
                                        If nValueType = ValueTypeEnum.Sum Then
                                            If (objFieldName.Sum = True And TableSection <> TableSectionEnum.Footer) Or (objFieldName.GrantTotalSum = True And TableSection = TableSectionEnum.Footer) Then
                                                If TableSection = TableSectionEnum.GroupFooter Then
                                                    sValue = "Sum(Fields!" & objFieldName.Key & ".Value," & Chr(34) & GroupName & Chr(34) & ")"
                                                Else
                                                    sValue = "Sum(Fields!" & objFieldName.Key & ".Value)"
                                                End If
                                            ElseIf TableSection = TableSectionEnum.Footer And objFieldName.Position = 1 And objFieldName.IsCurrency = False Then
                                                If bContainsGrandTotal = True Then
                                                    sValue = """Grand Total"""
                                                End If
                                            ElseIf TableSection = TableSectionEnum.Footer And ContainsGrouping = True And objFieldName.Position = 2 And objFieldName.IsCurrency = False Then
                                                If bContainsGrandTotal = True Then
                                                    sValue = """Grand Total"""
                                                End If
                                            End If
                                        Else
                                            sValue = "Fields!" & objFieldName.Key & ".Value"
                                        End If
                                        If sValue <> "" Then
                                            If objFieldName.IsCurrency = True Then
                                                sValue = "FormatNumber(" & sValue & ", 2, TriState.False, TriState.True, TriState.true)"
                                            ElseIf objFieldName.IsDate = True Then
                                                sValue = "IIF(" & sValue & " > """ & DateTime.MinValue & """,FormatDateTime(" & sValue & ", " & vbShortDate & "),"""")"
                                            End If
                                            If sValue <> "" Then sValue = "=" & sValue
                                        End If
                                    Case ValueTypeEnum.LiteralText
                                        If objFieldName.Position = 1 Then
                                            sValue = LiteralText
                                        End If
                                End Select
                            End If
                            .Value = sValue
                        End If
                    End With
                    Me.AddTextboxValue(WriterObject, objTxt, True)
                ElseIf bDataField = True Then
                    .WriteElementString("DataField", objFieldName.Key)
                End If
                Dim bFirst As Boolean = True
                For Each objE As String In objElementCol
                    .WriteEndElement()
                    If bFirst = True Then
                        bFirst = False
                        If nColspan > 0 Then 'And objE = "CellContents" Then
                            .WriteElementString("ColSpan", nColspan.ToString)
                        End If
                    End If
                Next objE
                If (TableSection = TableSectionEnum.FooterLiteralText Or TableSection = TableSectionEnum.GroupHeader) And objFieldName.Position = 1 Then
                    For i As Integer = 1 To objFields.Count - 1
                        .WriteElementString("TablixCell", "")
                    Next
                    Exit Do 'only one cell that spans the table
                End If
                objFieldName = Me.GetNextField(objFieldName)
            Loop
            .WriteEndElement()
            If bMembers = True Then
                .WriteEndElement() 'TablixMembers
            End If
        End With
    End Sub

    Private Function FieldsContainGrandTotal() As Boolean
        Dim b As Boolean = False
        For Each objF As bcField In objFields
            If objF.GrantTotalSum = True Then
                b = True
                Exit For
            End If
        Next
        Return b
    End Function
    Private Function FieldsContainSum() As Boolean
        Dim b As Boolean = False
        For Each objF As bcField In objFields
            If objF.Sum = True Then
                b = True
                Exit For
            End If
        Next
        Return b
    End Function
#Region "Add RDL"
    Private Sub AddTextboxValue(ByVal WriterObject As XmlWriter, ByVal BcTextboxObject As bcTextbox, BodyTextBox As Boolean) 'ByVal Value As String, ByVal FontFamilyValue As String, ByVal FontSizeValue As String, ByVal FontWeightValue As String, ByVal TextDecorationValue As String, ByVal ColorValue As String, ByVal TextAlignValue As String)
        '            <Paragraphs>
        '    <Paragraph>
        '        <TextRuns>
        '        <TextRun>
        '            <Value>Number</Value>
        '            <Style>
        '            <FontFamily>Verdana</FontFamily>
        '            <FontSize>8pt</FontSize>
        '            <FontWeight>Bold</FontWeight>
        '            <TextDecoration>Underline</TextDecoration>
        '            <Color></Color>
        '            </Style>
        '        </TextRun>
        '        </TextRuns>
        '        <Style>
        '        <TextAlign>Left</TextAlign>
        '        </Style>
        '    </Paragraph>
        '    </Paragraphs>
        '   <Style>
        '   <TopBorder>
        '        <Style>Solid</Style>
        '   </TopBorder>
        '   <VerticalAlign>Bottom</VerticalAlign>
        '   <PaddingRight>10pt</PaddingRight>
        '   </Style>
        '</Textbox>
        With WriterObject
            If BcTextboxObject.CanGrow <> "" Then .WriteElementString("CanGrow", "true")
            If BcTextboxObject.CanShrink <> "" Then .WriteElementString("CanShrink", "false")
            .WriteElementString("KeepTogether", "true")
            .WriteStartElement("Paragraphs")
            .WriteStartElement("Paragraph")
            .WriteStartElement("TextRuns")
            .WriteStartElement("TextRun")
            .WriteElementString("Value", BcTextboxObject.Value)
            'style
            .WriteStartElement("Style")

            If BcTextboxObject.FontFamily <> "" Then
                .WriteElementString("FontFamily", BcTextboxObject.FontFamily)
            End If
            If BcTextboxObject.FontSize <> "" Then
                .WriteElementString("FontSize", BcTextboxObject.FontSize)
            End If
            If BcTextboxObject.FontWeight <> "" Then
                .WriteElementString("FontWeight", BcTextboxObject.FontWeight)
            End If
            If BcTextboxObject.TextDecoration <> "" Then
                .WriteElementString("TextDecoration", BcTextboxObject.TextDecoration)
            End If
            If BcTextboxObject.Color <> "" Then
                .WriteElementString("Color", BcTextboxObject.Color)
            End If
            .WriteEndElement()
            'style
            .WriteEndElement() 'textrun
            .WriteEndElement() 'textruns
            .WriteStartElement("Style")
            If BcTextboxObject.TextAlign <> "" Then .WriteElementString("TextAlign", BcTextboxObject.TextAlign)

            .WriteEndElement() 'style
            .WriteEndElement() 'paragraph
            .WriteEndElement() 'paragraphs
            If BcTextboxObject.DefaultName <> "" Then
                .WriteElementString("DefaultName", "rd", BcTextboxObject.DefaultName)
            End If
            .WriteStartElement("Style") ' textbox style
            If BcTextboxObject.BorderAll = True Then
                Me.WriteBorder(WriterObject, "", BcTextboxObject.BorderWidth, BcTextboxObject.BorderColor, BcTextboxObject.BorderStyle)
            Else
                If BcTextboxObject.BorderTop = True Then Me.WriteBorder(WriterObject, "Top", BcTextboxObject.BorderTopWidth, BcTextboxObject.BorderColor, BcTextboxObject.BorderStyle)
                If BcTextboxObject.BorderBottom = True Then Me.WriteBorder(WriterObject, "Bottom", BcTextboxObject.BorderBottomWidth, BcTextboxObject.BorderColor, BcTextboxObject.BorderStyle)
                If BcTextboxObject.BorderRight = True Then Me.WriteBorder(WriterObject, "Right", BcTextboxObject.BorderWidth, BcTextboxObject.BorderColor, BcTextboxObject.BorderStyle)
                If BcTextboxObject.BorderLeft = True Then Me.WriteBorder(WriterObject, "Left", BcTextboxObject.BorderWidth, BcTextboxObject.BorderColor, BcTextboxObject.BorderStyle)
            End If
            If BcTextboxObject.PaddingTopWidth <> "" Then .WriteElementString("PaddingTop", BcTextboxObject.PaddingTopWidth)
            If BcTextboxObject.PaddingBottomWidth <> "" Then .WriteElementString("PaddingBottom", BcTextboxObject.PaddingBottomWidth)
            If BcTextboxObject.PaddingRightWidth <> "" Then .WriteElementString("PaddingRight", BcTextboxObject.PaddingRightWidth)
            If BcTextboxObject.PaddingLeftWidth <> "" Then .WriteElementString("PaddingLeft", BcTextboxObject.PaddingLeftWidth)
            If BcTextboxObject.VerticalAlign <> "" Then .WriteElementString("VerticalAlign", BcTextboxObject.VerticalAlign)
            If nReportType <> ReportTypeEnum.Excel And BodyTextBox = True And BcTextboxObject.AlternateBackColor = True Then
                .WriteElementString("BackgroundColor", "= IIf(RowNumber(Nothing) Mod 2 = 0, """ & "Silver" & """, ""Transparent"")")
            End If
            .WriteEndElement() 'textbox style
            If BcTextboxObject.Visible = False Then
                .WriteStartElement("Visibility")
                .WriteElementString("Hidden", "true")
                .WriteEndElement()
            End If
        End With
    End Sub
    Private Sub WriteBorder(ByVal WriterObject As XmlWriter, ByVal Dimension As String, ByVal WidthValue As String, ByVal ColorValue As String, ByVal StyleValue As String)
        If ColorValue = "" Then ColorValue = "Black"
        If StyleValue = "" Then StyleValue = "Solid"
        With WriterObject
            .WriteStartElement(Dimension & "Border")
            .WriteElementString("Style", StyleValue)
            .WriteElementString("Width", WidthValue)
            .WriteElementString("Color", ColorValue)

            .WriteEndElement()
        End With
    End Sub
    Private Sub AddSectionLine(ByVal WriterObject As XmlWriter, ByVal Position As LinePositionEnum)
        With WriterObject
            .WriteStartElement(Position.ToString & "Border")
            .WriteElementString("Style", "Solid")
            .WriteEndElement()
        End With
    End Sub
    Private Sub AddSectionPadding(ByVal WriterObject As XmlWriter, Optional ByVal TopValue As String = "2pt", Optional ByVal BottomValue As String = "2pt")
        With WriterObject
            .WriteElementString("PaddingTop", TopValue)
            .WriteElementString("PaddingBottom", BottomValue)
        End With
    End Sub

#End Region


#Region "Misc"
    Private Function GetDataToString(ByVal TableObj As DataTable) As String
        Dim sData As String = ""
        Dim nPos As Integer = 0
        Dim sHeader As String = ""
        Dim objHeaderFieldName As bcField = Me.GetNextField(Nothing)
        Do Until objHeaderFieldName Is Nothing = True
            With objHeaderFieldName
                If .Visible = True Then
                    sHeader = sHeader & .Header & vbTab
                End If
            End With
            objHeaderFieldName = Me.GetNextField(objHeaderFieldName)
        Loop
        For Each objR As DataRow In TableObj.Rows
            Dim objFieldName As bcField = Me.GetNextField(Nothing)
            Dim sRow As String = ""
            Do Until objFieldName Is Nothing = True
                nPos = objFieldName.Position
                With objFieldName
                    If .Visible = True Then
                        sRow = sRow & CStr(objR.Item(.Key) & "") & vbTab
                    End If
                End With
                objFieldName = Me.GetNextField(objFieldName)
            Loop
            If sRow <> "" Then sData = sData & vbCrLf & sRow
        Next objR
        sData = sHeader & vbCrLf & sData
        Return sData
    End Function
    Private Sub SetFieldWidths()
        For Each objF As bcField In objFields
            With objF
                If objF.IsDate = True And objF.Width <> 0 And objF.CustomWidth <> -1 Then
                    'keep as is 
                ElseIf objF.CustomWidth <> 0 Then
                    .Width = .CustomWidth
                ElseIf .Width > 1.5 Then
                    .Width = 1.7
                ElseIf objF.Width < 0.6 Then
                    .Width = 0.5
                Else
                    .Width = 0.75
                End If
            End With
            Debug.Print(objF.Header & " " & objF.Width)
        Next
    End Sub
    Private Sub SetPaperSize()
        Select Case System.Globalization.CultureInfo.CurrentCulture.EnglishName
            Case "English (United States)"
                nPageWidth = 8.5
                nPageHeight = 11
            Case Else
                nPageWidth = 8.27
                nPageHeight = 11.69
        End Select
    End Sub
    Private Sub AddFieldsFromBindingString(ByVal BindingString As String, ByVal ExcludeList As Collection)
        Dim sGridBinding() As String = BindingString.Split(",")
        Dim nPos As Integer = 1
        For Each sBinding As String In sGridBinding
            Dim s() As String = sBinding.Split(":")
            Dim objF As bcField = Nothing
            Dim bExclude As Boolean = False
            If ExcludeList Is Nothing = False Then
                If ExcludeList.Contains(s(0)) = True Then
                    bExclude = True
                End If
            End If
            If bExclude = False Then
                Dim sField As String = s(0)
                Dim sHeader As String = ""
                Dim nWidth As Decimal = 0
                Dim nPadding As Decimal = 6 'where using the default that reporting service uses
                If UBound(s) >= 1 Then
                    sHeader = s(1)
                End If
                If UBound(s) >= 2 Then
                    If IsNumeric(s(2)) = True Then
                        nWidth = Decimal.Parse(s(2), objFormat)
                    End If
                End If
                If UBound(s) >= 3 Then
                    If IsNumeric(s(3)) = True Then
                        nPadding = Decimal.Parse(s(3), objFormat)
                    End If
                End If
                Me.AddBindingField(sField, sHeader, nWidth, nPadding, nPos)
                nPos = nPos + 1
            End If
        Next
    End Sub
    Private Sub DoFormatGLAccountSegments()
        ' Comp Loc Dept Sect Acct 
        Dim nWidth As Double = 0.45
        Try
            Dim objComp As bcField = objFields("cCompanyCode")
            Dim objLoc As bcField = objFields("cLocationCode")
            Dim objDept As bcField = objFields("cDepartmentCode")
            Dim objSect As bcField = objFields("cGLSectionCode")
            Dim objAcct As bcField = objFields("cAccountCode")
            If objComp Is Nothing = False And objLoc Is Nothing = False And objSect Is Nothing = False And objAcct Is Nothing = False Then
                objComp.CustomWidth = nWidth
                objLoc.CustomWidth = nWidth
                objDept.CustomWidth = nWidth
                objSect.CustomWidth = nWidth
                objAcct.CustomWidth = nWidth
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Sub HideBindingField(ByVal DataField)
        Me.AddBindingField(DataField, "", 0, 0, False)
    End Sub
    Public Sub AddBindingField(ByVal DataField As String, ByVal Header As String, ByVal Width As Decimal, ByVal Padding As Decimal, ByVal Position As Integer, Optional ByVal Visible As Boolean = True)
        Dim objF As bcField
        If objFields.Contains(DataField) Then
            objF = objFields(DataField)
        Else
            objF = New bcField
            objF.Key = DataField
        End If
        With objF
            .Header = Header
            .CustomWidth = Width
            .CustomPosition = Position
            .Padding = Padding
            .Visible = Visible
        End With
        If objFields.Contains(DataField) = False Then
            objFields.Add(objF, objF.Key)
        End If
    End Sub
    Private Sub SetPageOrientation()
        Me.SetPaperSize()
        If Me.PageOrientation = PageOrientationEnum.NotSet Then
            Dim nTotal As Double = 0
            For Each objF As bcField In objFields
                With objF
                    If .Visible = True Then
                        If .Width > -1 Then
                            nTotal = nTotal + .Width
                        End If
                    End If
                End With
            Next
            Debug.Print("Total width = " & nTotal.ToString & " Report Width = " & Me.ReportWidth.ToString)
            If nTotal > (Me.ReportWidth) Then
                Me.PageOrientation = PageOrientationEnum.Landscape
            End If
        End If
    End Sub
    Private Function GetNextField(ByVal PreviousField As bcField) As bcField
        Dim nNextPos As Integer = 0
        Dim nPreviousPos As Integer
        Dim objField As bcField = Nothing
        Dim sPreviousKey As String
        If PreviousField Is Nothing Then
            nPreviousPos = 0
            sPreviousKey = ""
        Else
            nPreviousPos = PreviousField.Position
            sPreviousKey = PreviousField.Key
        End If
        For Each objF As bcField In objFields
            Dim nPos As Integer = objF.Position
            Dim nCustomPos As Integer = objF.CustomPosition
            If objF.Key <> sPreviousKey Then
                If nNextPos = 0 And nPos > nPreviousPos Then 'first one in loop
                    nNextPos = nPos
                    objField = objF
                ElseIf nCustomPos > nPreviousPos And nCustomPos < nNextPos Then
                    nNextPos = nPos
                    objField = objF
                ElseIf nPos > nPreviousPos And nPos < nNextPos Then
                    nNextPos = nPos
                    objField = objF
                End If
            End If
        Next
        Return objField
    End Function

    Public Shared Function CurrencyFormat(ByVal IncludeNegativeFormat As Boolean) As String
        Dim sCurrencyFormat As String = ""
        sCurrencyFormat = "n2"
        Return sCurrencyFormat
    End Function

    Public Function GetEnglishNameForField(ByVal Value As String) As String
        Dim sName As String = Value
        Dim sProper As String = ""
        If sName.Contains(PivotFieldDelimiter) Then
            Dim sFields() As String = Split(sName, PivotFieldDelimiter)
            sName = ""
            For Each sField As String In sFields
                If sField <> "" Then
                    Dim sPart As String = GetEnglishNameForField(sField)
                    sName = sName & sPart & " "
                End If
            Next
            If sName.Length > 0 Then
                sName = Trim(sName)
            End If
        End If
        If sName.StartsWith("i") = True Or sName.StartsWith("m") = True Or sName.StartsWith("b") = True Then
            sName = sName.Substring(1)
        ElseIf sName.StartsWith("dt") = True Then
            sName = sName.Substring(2)
        ElseIf sName.StartsWith("dt") = True Or sName.StartsWith("vch") = True Then
            sName = sName.Substring(3)
        ElseIf sName.StartsWith("c") And sName.Substring(1, 1).ToLower <> sName.Substring(1, 1) Then 'first letter is c and second letter is upper case
            sName = sName.Substring(1)
        End If
        If sName.ToLower.EndsWith("void") = False And sName.ToLower.EndsWith("id") = True And sName.ToLower.EndsWith(" id") = False Then sName = sName.Substring(0, sName.Length - 2)
        Dim n As Integer = 0
        Dim sRemaining As String = ""
        For i As Integer = 0 To sName.Length - 1
            Dim sChar As String = sName(i)
            If sChar <> sChar.ToLower Then
                If sProper <> "" Then
                    If StrReverse(sProper).Substring(0, 1) = StrReverse(sProper).Substring(0, 1).ToLower Then 'last char is lowercase
                        sProper = sProper & " "
                    End If
                End If
                If i - n > 0 Then
                    sProper = sProper & sName.Substring(n, i - n)
                    Dim nRemaining As Integer = i
                    If nRemaining < sName.Length Then sRemaining = sName.Substring(nRemaining)
                End If
                n = i
            End If
        Next
        If sProper <> "" Then sProper = sProper & " " & sRemaining
        If sProper <> "" Then sName = sProper
        Return sName
    End Function
    Private Function GetGrandTotalsFromTable(ByVal TableObj As DataTable) As Collection
        Dim objCol As Collection = New Collection
        For Each objC As DataColumn In TableObj.Columns
            If objC.DataType Is GetType(Decimal) Or objC.DataType Is GetType(Double) Then
                Dim objG As bcGrandTotal = New bcGrandTotal
                objG.Datafield = objC.ColumnName
                objG.Caption = objC.ColumnName
                objCol.Add(objG, objC.ColumnName)
            End If
        Next
        Return objCol
    End Function

#End Region

#Region "Properties"
    Private Function ReportWidth() As Double
        Return nPageWidth - (2 * nSideMargin) '- 0.5
    End Function
    Private Function ReportHeight() As Double
        Return 2 'nPageHeight - nTopMargin - nBottomMargin
    End Function
    Public ReadOnly Property DataToString()
        Get
            Return sDataToString
        End Get
    End Property

    Private Function GetSegmentValue(ByVal FieldName As String, ByVal CollectionObj As Collection) As Boolean
        Dim b As Boolean = True
        If CollectionObj.Count > 0 Then
            If CollectionObj.Contains(FieldName) = False Then
                If CollectionObj.Contains(FieldName) = False Then
                    b = False
                Else
                    Dim sValue As String = CollectionObj(FieldName)
                    Dim s() As String = sValue.Split(":")
                    If s.Length > 1 Then
                        If s(1) = "0" Then
                            b = False
                        End If
                    End If
                End If
            End If
        End If
        Return b
    End Function
    Private Function GetLineCount(ByVal Value As String) As Integer
        Dim n As Integer = 1
        For Each s As Char In Value.Trim.ToCharArray
            If s = vbCr Then n = n + 1
            If s = vbLf Then n = n + 1
        Next
        Return n
    End Function
#End Region


    Private Class bcField
        Public Width As Decimal = 0
        Public TextAlignment As String = ""
        Public Header As String = ""
        Public Key As String = ""
        Public IsCurrency As Boolean = False
        Public IsDate As Boolean = False
        Public Sum As Boolean
        Public GrantTotalSum As Boolean
        Public Position As Integer
        Public CustomWidth As Decimal = 0
        Public CustomPosition As Integer = 0
        Public Visible As Boolean = True
        Public Padding As Decimal = 0
    End Class

    Private Class bcTextbox
        Public Value As String
        Public FontFamily As String
        Public FontSize As String
        Public FontWeight As String
        Public TextDecoration As String
        Public Color As String
        Public TextAlign As String
        Public VerticalAlign As String
        Public BorderTop As Boolean
        Public BorderBottom As Boolean
        Public BorderRight As Boolean
        Public BorderLeft As Boolean
        Public BorderAll As Boolean
        Public BorderWidth As String
        Public BorderColor As String
        Public AlternateBackColor As Boolean
        Public BorderStyle As String
        Public PaddingTopWidth As String
        Public PaddingBottomWidth As String
        Public PaddingLeftWidth As String
        Public PaddingRightWidth As String
        Public CanGrow As String
        Public CanShrink As String
        Public DefaultName As String
        Dim sBorderTopWidth As String
        Dim sBorderBottomWidth As String
        Public Visible As Boolean = True
        Public Sub New()
            MyBase.New()
        End Sub
        Public Sub New(ByVal TextValue As String, ByVal FontFamilyValue As String, ByVal FontSizeValue As String, ByVal FontWeightValue As String, ByVal TextDecorationValue As String, ByVal ColorValue As String, ByVal TextAlignValue As String)
            MyBase.New()
            Value = TextValue
            FontFamily = FontFamilyValue
            FontSize = FontSizeValue
            FontWeight = FontWeightValue
            TextDecoration = TextDecorationValue
            Color = ColorValue
            TextAlignValue = TextAlign
        End Sub
        Public Property BorderTopWidth() As String
            Set(ByVal Value As String)
                sBorderTopWidth = Value
            End Set
            Get
                Dim s As String = sBorderTopWidth
                If sBorderTopWidth = "" And BorderWidth <> "" Then
                    sBorderTopWidth = BorderWidth
                End If
                Return s
            End Get
        End Property

        Public Property BorderBottomWidth() As String
            Set(ByVal Value As String)
                sBorderBottomWidth = Value
            End Set
            Get
                Dim s As String = sBorderBottomWidth
                If sBorderBottomWidth = "" And BorderWidth <> "" Then
                    sBorderBottomWidth = BorderWidth
                End If
                Return s
            End Get
        End Property
    End Class

    Public Sub New()
        MyBase.New()
        objFormat.NumberDecimalSeparator = "."
        objFormat.NumberGroupSeparator = ","
    End Sub
    Public Function CurrencyFormat() As String
        Dim sCurrencyFormat As String = ""
        sCurrencyFormat = "n2"
        Return sCurrencyFormat
    End Function
    Public Function NumericFormat(ByVal DecimalPlaces As Integer) As String
        Dim sNumericFormat As String = ""
        Dim ZeroString As String = New String("0", DecimalPlaces)
        If sNumericFormat = "" Then
            Dim sDecimal As String = Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
            Dim sComma As String = Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyGroupSeparator
            Dim sFormat As String = "#" & sComma & "0" & sDecimal & ZeroString
            sNumericFormat = sFormat & ";(" & sFormat & ")"
        End If
        Return sNumericFormat
    End Function
    Public Sub LoadReportFromTable(ReportViewerObj As ReportViewer, ByVal ReportHeaderValue As String, ByVal TableObj As DataTable, Optional BindingString As String = "")
        Dim ReportType As RdlGenerator.ReportTypeEnum
        Dim ExcludeList As Collection = Nothing
        Dim DefaultToLandscape As Boolean
        Dim IncludeCriteria As Boolean

        Dim objRDL As RdlGenerator = New RdlGenerator()
        Dim sReportHeaderSource As String
        Dim objExcludeListSource As Collection
        Dim bDefaultToLandscapeSource As Boolean
        Dim objAdapterSource As DataTable
        Dim bIncludeCriteria As Boolean = False
        Dim objReportViewer As ReportViewer = ReportViewerObj
        Try
            objReportViewer.Clear()
            objReportViewer.Reset()
        Catch ex As Exception
        End Try
        Dim sGroupField As String = ""
        Dim sTotalCols As String = ""
        Dim sGrandTotalCols As String = ""
        Dim sCriteria As String = ""
        Dim sData As String = ""
        Dim sExtraHeader As String = ""
        sReportHeaderSource = ReportHeaderValue
        objExcludeListSource = ExcludeList
        bDefaultToLandscapeSource = DefaultToLandscape
        bIncludeCriteria = IncludeCriteria
        objAdapterSource = TableObj

        Dim TableObject As DataTable = Nothing
        If TableObj Is Nothing = False Then
            TableObject = TableObj
        End If
        Dim objMemoryStream As IO.MemoryStream
        Dim objTable As DataTable
        Dim objLocalReport As LocalReport = objReportViewer.LocalReport()
        Dim objReportDS As ReportDataSource = New ReportDataSource
        'Me.SetWindowTitle()
        objReportViewer.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local

        If TableObject Is Nothing = True Then
            Throw New Exception("Datasource for report not specified")
        End If

        With TableObject
            objTable = .Clone
            For Each objRowView As DataRowView In .DefaultView
                objTable.ImportRow(objRowView.Row)
            Next
            objTable.DefaultView.Sort = .DefaultView.Sort
        End With
        Dim sLocalCriteria As String = ""
        'sCriteria = modGlobal.GetCriteriaToString(tblCriteria, tblCriteria)

        '    If IncludeCriteria = True Then
        '        sLocalCriteria = sCriteria
        '    End If
        objMemoryStream = objRDL.Run(sReportHeaderSource, objTable, BindingString, RdlGenerator.ReportTypeEnum.Preview, ExcludeList, sLocalCriteria, sGroupField, sExtraHeader, sTotalCols, sGrandTotalCols)
        sData = objRDL.DataToString


        If TableObj Is Nothing = False Then
            Dim enc As New System.Text.UTF8Encoding
            Dim arrBytData() As Byte = New Byte(CType(objMemoryStream.Length, Long)) {}
            Dim s As String = ""
            objMemoryStream.Read(arrBytData, 1, objMemoryStream.Length)
            s = enc.GetString(arrBytData, 1, arrBytData.Length - 1)
            Dim objMemoryStream2 As System.IO.MemoryStream

            s = s.Replace("ReplaceExtraHeader", sExtraHeader)
            objMemoryStream2 = New System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(s))

            objLocalReport.LoadReportDefinition(objMemoryStream2)
        Else
            'Dim enc As New System.Text.UTF8Encoding
            'Dim arrBytData() As Byte = New Byte(CType(objMemoryStream.Length, Long)) {}
            'Dim s As String = ""
            'objMemoryStream.Read(arrBytData, 1, objMemoryStream.Length)
            's = enc.GetString(arrBytData, 1, arrBytData.Length - 1)
            objLocalReport.LoadReportDefinition(objMemoryStream)
        End If

        With objReportDS
            .Name = "DataSet1"
            .Value = objTable
        End With
        objLocalReport.DataSources.Add(objReportDS)
        objReportViewer.RefreshReport()
        Select Case ReportType
            'Case RdlGenerator.ReportTypeEnum.Preview
            '    Me.Show()
            Case RdlGenerator.ReportTypeEnum.Excel
                'ExportToExcel()
        End Select
        objMemoryStream.Close()
        objMemoryStream = Nothing
        objRDL = New RdlGenerator 'ensure that next time a new instance is used, we can't set this at the bginning of the procedure because the instance may have properties set by an event in bcGrid
    End Sub

End Class 'RdlGenerator


