Option Explicit On
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesFile

Public Class Nova_projekceButton2

    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Protected Overrides Sub OnClick()
         Try
            'after click on "New projection" button, create new Data Frame, set its size, color, background, add shapefile with diagram outline 
            Dim pMxDoc As IMxDocument = Nothing
            pMxDoc = My.ArcMap.Document

            'Create a new map
            Dim pMap As IMap = Nothing
            pMap = New Map
            pMap.Name = "Equal-area projection, lower hemisphere"

            'Create a new MapFrame and associate map with it
            Dim pMapFrame As IMapFrame = Nothing
            pMapFrame = New MapFrame
            pMapFrame.Map = pMap

            'Set the position of the new map frame
            Dim pElement As IElement = Nothing
            Dim pEnv As IEnvelope = Nothing
            pElement = pMapFrame
            pEnv = New Envelope
            pEnv.PutCoords(0, 0, 5, 5)
            pElement.Geometry = pEnv

            'Add mapframe to the layout
            Dim pGraphicsContainer As IGraphicsContainer = Nothing
            pGraphicsContainer = pMxDoc.PageLayout
            pGraphicsContainer.AddElement(pMapFrame, 0)

            'Make the newly added map the focus map
            Dim pActiveView As IActiveView
            pActiveView = pMxDoc.ActiveView
            If TypeOf pActiveView Is IPageLayout Then
                pActiveView.FocusMap = pMap
            Else
                pMxDoc.ActiveView = pMap
            End If


            'background color of the new Data Framu
            Dim pFillSymbol As IFillSymbol
            Dim pRgbColor As IRgbColor
            Dim pSymbolBackground As ISymbolBackground

            'Get a reference to the layout's graphics container
            'Set pMxDoc = Application.Document
            'Set pLayoutGraphicsContainer = pMxDoc.PageLayout

            'Find the map frame associated with the focus map
            'Set pMapFrame = pLayoutGraphicsContainer.FindFrame(pMxDoc.FocusMap)
            'If pMapFrame Is Nothing Then Exit Sub

            'Associate a SymbolBackground with the frame
            pSymbolBackground = pMapFrame.Background
            'If pSymbolBackground Is Nothing Then
            pSymbolBackground = New SymbolBackground
            'End If
            pFillSymbol = New SimpleFillSymbol
            pRgbColor = New RgbColor
            pRgbColor.Red = 255
            pRgbColor.Green = 255
            pRgbColor.Blue = 255
            pFillSymbol.Color = pRgbColor
            pSymbolBackground.FillSymbol = pFillSymbol
            pMapFrame.Background = pSymbolBackground

            'Refresh ActiveView and TOC
            pActiveView.Refresh()
            pMxDoc.CurrentContentsView.Refresh(0)


            'Get Desktop Adress
            Dim filePath As String
            filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

            Dim a As String = Nothing
            Dim fileContents As String
            'fileContents = My.Computer.FileSystem.ReadAllText("C:\temp\TableDef.txt")
            fileContents = My.Computer.FileSystem.ReadAllText(filePath & "\stereonet.txt")
            Dim b As String = Nothing
            Using MyReader As New  _
            Microsoft.VisualBasic.FileIO.TextFieldParser(filePath & "\stereonet.txt")

                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim CurrentRow As String()

                While Not MyReader.EndOfData
                    CurrentRow = MyReader.ReadFields()
                    a = CurrentRow(0)
                End While
            End Using

            'add shapefile with outline (cross.shp) focus, refresh TOC
            Dim pworkspaceFactory As IWorkspaceFactory
            Dim pfeatureWorkspace As IFeatureWorkspace
            Dim pFeatureLayer As IFeatureLayer
            'Create a new ShapefileWorkspaceFactory object and open a shapefile folder
            pworkspaceFactory = New ShapefileWorkspaceFactory
            pfeatureWorkspace = pworkspaceFactory.OpenFromFile(a, 0)
            'Create a new FeatureLayer and assign a shapefile to it
            pFeatureLayer = New FeatureLayer
            pFeatureLayer.FeatureClass = pfeatureWorkspace.OpenFeatureClass("cross.shp")
            pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
            pFeatureLayer.ScaleSymbols = True
            'Add the FeatureLayer to the focus map
            pMap.AddLayer(pFeatureLayer)
            pMxDoc.UpdateContents()

            'switch to Layout View, if it is not already
            If Not TypeOf pMxDoc.ActiveView Is IPageLayout Then
                'Set pMxDoc.ActiveView = pMxDoc.FocusMap
                'Else
                pMxDoc.ActiveView = pMxDoc.PageLayout
            End If

            'Refresh the map frame
            pActiveView = pMxDoc.FocusMap
            pActiveView.PartialRefresh(esriViewDrawPhase.esriViewBackground, Nothing, Nothing)

        Catch ex As Exception
            'for the first time, OpenFileForm is loaded with info and file browser
            'pokud nenačte již uložený txt s cestou k shp, otevře openfile form,který by měl 
            'přidávat shp do mapy a popisuje, že jde o shp stažený s mxd z netu
            Dim frm2 As New OpenFileForm
            frm2.Show()
        End Try
    End Sub

End Class


