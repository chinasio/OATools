Public Class LinkBackButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        'open form
        Dim frmZ As New LinkBackForm
        frmZ.Show()

  End Sub

    'Protected Overrides Sub OnUpdate()

    'End Sub
End Class
