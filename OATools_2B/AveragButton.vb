Public Class AveragButton
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        'open form
        Dim frmSA As New AveragForm
        frmSA.Show()

    End Sub

    'Protected Overrides Sub OnUpdate()

    'End Sub
End Class
