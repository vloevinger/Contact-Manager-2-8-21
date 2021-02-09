Public Class bizReport
    Inherits bizObject(Of ParamEnum)

    Public Enum ParamEnum
        iClientEventId
        vchReportType
    End Enum

    Public Sub New()
        MyBase.New("")
    End Sub

    Public Overridable Function RunReport(ReportTypeVal As String, Optional ClientEventId As Long = 0) As DataTable
        Dim objT As DataTable = Nothing

        Using objCmd As SqlCommand = mdUtility.GetSQLCommand("ReportGet")
            Me.SetCommandParamValue(objCmd, ParamEnum.vchReportType, ReportTypeVal)
            Me.SetCommandParamValue(objCmd, ParamEnum.iClientEventId, ClientEventId)
            objT = mdUtility.ExecuteSQL(objCmd)
        End Using

        Return objT

    End Function
End Class
