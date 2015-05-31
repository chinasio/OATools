Public Class Density_Button
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim frmW As New DensityForm
        frmW.Show()
  End Sub

    'Protected Overrides Sub OnUpdate()

    'End Sub
End Class
