Public Class FoldAxesButton
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        'open form
        Dim frmFA As New FoldAxesForm
        frmFA.Show()

    End Sub

End Class
