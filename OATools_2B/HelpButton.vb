Public Class HelpButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        'open form
        Dim frmHelp As New HelpForm
        frmHelp.Show()
  End Sub

    ' Protected Overrides Sub OnUpdate()

    ' End Sub
End Class
