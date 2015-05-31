Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.ArcMapUI

Public Class OpenFileForm

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            'folder browser
            Dim FolderBrowserDialog1 As New FolderBrowserDialog
            With FolderBrowserDialog1

                '.RootFolder = Environment.SpecialFolder.MyDocuments
                .SelectedPath = "c:"
                .Description = "Select folder"

                If .ShowDialog = DialogResult.OK Then
                    TextBox1.Text = (.SelectedPath)
                End If
            End With

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'add cross.shp from selected path
        'read the map
        Dim pMxDoc As IMxDocument
        pMxDoc = My.ArcMap.Document
        Dim pMap As IMap
        pMap = pMxDoc.FocusMap

        'add shapefile, focus, refresh TOC, switch to Layout view
        Dim pworkspaceFactory As IWorkspaceFactory
        Dim pfeatureWorkspace As IFeatureWorkspace
        Dim pFeatureLayer As IFeatureLayer
        'Create a new ShapefileWorkspaceFactory object and open a shapefile folder
        pworkspaceFactory = New ShapefileWorkspaceFactory
        pfeatureWorkspace = pworkspaceFactory.OpenFromFile(TextBox1.Text, 0)
        'Create a new FeatureLayer and assign a shapefile to it
        pFeatureLayer = New FeatureLayer
        pFeatureLayer.FeatureClass = pfeatureWorkspace.OpenFeatureClass("cross.shp")
        pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
        pFeatureLayer.ScaleSymbols = True
        'Add the FeatureLayer to the focus map
        pMap.AddLayer(pFeatureLayer)
        pMxDoc.UpdateContents()

        'write path to text file
        Try
            Dim filePAth As String
            filePAth = CreateObject("WScript.Shell").SpecialFolders("Desktop")
            My.Computer.FileSystem.WriteAllText(filePAth & "\stereonet.txt", (TextBox1.Text), False)
        Catch ex As Exception
            Exit Sub
        End Try

        'write path to text file
        'Try
        'My.Computer.FileSystem.WriteAllText("C:\temp\stereonet.txt", (TextBox1.Text), False)
        'Catch ex As Exception
        'Exit Sub
        'End Try



    End Sub

    Private Sub OpenFileForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class