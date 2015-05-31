Public Class PlotTool
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        'open form
        Try
            Dim frm1 As New PlotForm
            frm1.Show()
        Catch ex As Exception
            MsgBox("Show PlotForm error. " & Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub
End Class
