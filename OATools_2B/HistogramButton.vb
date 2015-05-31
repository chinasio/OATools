Public Class HistogramButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        'open form
        Dim frmS As New HistogramForm
        frmS.show()
  End Sub

    'Protected Overrides Sub OnUpdate()

    'End Sub
End Class
