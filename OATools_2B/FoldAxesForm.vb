Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Math
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.esriSystem

Public Class FoldAxesForm

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            'after the input layer is chosen
            'read the layer
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap
            Dim pFLDB As IFeatureLayer

            Dim aLName As String
            Dim I As Long

            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            'input layer must't be a raster layer
            If TypeOf pMap.Layer(I) Is IRasterLayer Then
                MsgBox("You´ve chosen a raster layer.. but we need vector point feature class or shapefile!")
                Exit Sub

            End If

            pFLDB = pMap.Layer(I)
            'input layer must have data source
            If pFLDB.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            'input layer mustn't be of a sde data source
            If pFLDB.DataSourceType = "SDE Feature Class" Then
                MsgBox("Input layer is a SDE Feature Class, which is not editable. Export it to shapefile or database Feature Class first.")
                Exit Sub
            End If

            Dim pFClassDB As IFeatureClass
            pFClassDB = pFLDB.FeatureClass

            Dim pTableDB As ITable
            pTableDB = pFClassDB
            Dim pFieldsDB As IFields
            pFieldsDB = pTableDB.Fields
            'Dim pFieldDB As IField

            'read value to comboboxes
            ComboBox2.Items.Add("")
            ComboBox3.Items.Add("")

            'read field names to comboboxes
            Dim j As Long
            For j = 0 To pFieldsDB.FieldCount - 1
                If pFieldsDB.Field(j).Type = esriFieldType.esriFieldTypeDouble Or pFieldsDB.Field(j).Type = esriFieldType.esriFieldTypeInteger Then
                    ComboBox2.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox3.Items.Add(pFieldsDB.Field(j).Name)
                End If
            Next

            'fill the field names for dipdir and dip automatically if they have "smer_s" and "sklon_s" names'Dim W As Long
            Dim y As Long
            Dim z As Long

            y = pTableDB.FindField("DIPDIR")
            z = pTableDB.FindField("DIP")

            If y > (-1) Then
                ComboBox2.Text = "DIPDIR"
            End If
            If z > (-1) Then
                ComboBox3.Text = "DIP"
            End If

            'output name
            TextBox2.Text = "Axes_" & ComboBox1.Text

            'display number of selected features
            Dim pFSelection As IFeatureSelection
            pFSelection = pFLDB
            Dim pFLayerDB As ISelectionSet
            pFLayerDB = pFSelection.SelectionSet

            Label16.Text = pFLayerDB.Count & " features selected"

            'read input layer and learn its extent
            Dim pEnvelope As IEnvelope = Nothing
            Dim MinX, MaxX, MinY, MaxY As Double

            If pFLayerDB.Count > 0 Then
                Dim pEnumGeom As IEnumGeometry
                Dim pEnumGeomBind As IEnumGeometryBind
                pEnumGeom = New EnumFeatureGeometry
                pEnumGeomBind = pEnumGeom
                pEnumGeomBind.BindGeometrySource(Nothing, pFLayerDB)

                Dim pGeomFactory As IGeometryFactory
                pGeomFactory = New GeometryEnvironment

                Dim pGeom As IGeometry
                pGeom = pGeomFactory.CreateGeometryFromEnumerator(pEnumGeom)

                pEnvelope = pGeom.Envelope
            Else
                Dim pGeoDataset As IGeoDataset
                Dim pSpatialRef As ISpatialReference
                pGeoDataset = pFLDB  'QI
                pSpatialRef = pGeoDataset.SpatialReference
                pEnvelope = pGeoDataset.Extent
            End If

            'read radius of influence
            Dim pv As Double
            pv = TextBox4.Text
            'read the interval of averaging stations
            Dim n As Double
            n = TextBox3.Text

            MinX = pEnvelope.XMin
            MinY = pEnvelope.YMin
            MaxX = pEnvelope.XMax
            MaxY = pEnvelope.YMax
            MinX = MinX - n
            MinY = MinY - n
            MaxX = MaxX + n
            MaxY = MaxY + n

            'calculate extent in X and Y direction and number of stations
            Dim dX, dY As Double
            dX = MaxX - MinX
            dY = MaxY - MinY

            'display extent and number of stations in the form
            TextBox7.Text = Round(dX, 0)
            TextBox8.Text = Round(dY, 0)
            TextBox5.Text = Round(TextBox7.Text / TextBox3.Text, 0) + 1
            TextBox6.Text = Round(TextBox8.Text / TextBox3.Text, 0) + 1

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            'After click on "Calculate fold axes" button
            'Dim dipdirF As String
            'dipdirF = ComboBox2.Text

            'input layer must be chosen
            If ComboBox1.Text = "" Then
                MsgBox("Choose an input layer!")
                Exit Sub
            End If
            'fields for dipdir and dip must be selected
            If ComboBox2.Text = "" Or ComboBox3.Text = "" Then
                MsgBox("Choose fields!")
                Exit Sub
            End If
            'workspace must be set
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Select workspace, output name, interval and radius of influence!")
                Exit Sub
            End If

            'read the input layer
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap
            'Dim pLDB As ILayer
            Dim pFLDB As IFeatureLayer = Nothing

            Dim aLName As String
            Dim I As Long

            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            pFLDB = pMap.Layer(I)
            'input layer must have data source
            If pFLDB.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            'read the selection set
            Dim pFSelection As IFeatureSelection = Nothing
            pFSelection = pFLDB
            Dim pFLayerDB As ISelectionSet = Nothing
            pFLayerDB = pFSelection.SelectionSet

           'pFLayerDB is a selectionset, if no selection is made, all features are read
            If pFLayerDB.Count = 0 Then
                pFSelection.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                pFLayerDB = pFSelection.SelectionSet

                'clear selection
                pFSelection.Clear()
            End If

            'extent of selected features
            Dim pEnvelope As IEnvelope = Nothing
            Dim pEnumGeom As IEnumGeometry = Nothing
            Dim pEnumGeomBind As IEnumGeometryBind = Nothing
            pEnumGeom = New EnumFeatureGeometry
            pEnumGeomBind = pEnumGeom
            pEnumGeomBind.BindGeometrySource(Nothing, pFLayerDB)

            Dim pGeomFactory As IGeometryFactory = Nothing
            pGeomFactory = New GeometryEnvironment

            Dim pGeom As IGeometry = Nothing
            pGeom = pGeomFactory.CreateGeometryFromEnumerator(pEnumGeom)
            Dim MinX, MaxX, MinY, MaxY As Double
            pEnvelope = pGeom.Envelope

            'read the radius of inluence from the form, textbox4, set by user.
            Dim pv As Double = Nothing
            pv = TextBox4.Text

            'read the interval of averaging stations
            Dim n As Double = Nothing
            n = TextBox3.Text

            MinX = pEnvelope.XMin
            MinY = pEnvelope.YMin
            MaxX = pEnvelope.XMax
            MaxY = pEnvelope.YMax
            'extent
            MinX = MinX - n
            MinY = MinY - n
            MaxX = MaxX + n
            MaxY = MaxY + n

            'etxent in X direction
            Dim dX As Double = Nothing
            dX = MaxX - MinX

            'number of stations in Xdirection
            Dim nX As Double = Nothing
            nX = Round(dX / n, 0) + 1

            'extent must´t be enormaly large, if so, varnend exit sub
            'this is because of the mistaken or missing  coordinates in the input table.
            If (MaxX - MinX) * (MaxY - MinY) > 250000000000.0# Then
                MsgBox("Area of selected data is larger than 250 000km2 (" & (MaxX - MinX) * (MaxY - MinY) / 1000000 & " km2, that´s too much!." & _
                "Maybe you have a mistake in coordinates. E.g. one digit is more or less. Check it!")
                Exit Sub
            End If

            'new shapefile is created using Geoprocessingtool
            'read the spatial reference from the input layer
            Dim pSpatialReference As ISpatialReference = Nothing
            Dim geoDataset As IGeoDataset = Nothing
            geoDataset = pFLDB
            pSpatialReference = geoDataset.SpatialReference

            'call the Geoprocessing tool
            Dim pGP As IGeoProcessor = Nothing
            pGP = New GeoProcessor
            pGP.OverwriteOutput = True


            'create parameters array for the "CreateFeatureClass" tool
            Dim pParameterArray As IVariantArray = Nothing
            pParameterArray = New VarArray

            'workspace path is read from the form
            Dim filePath, name As String
            filePath = TextBox1.Text
            name = TextBox2.Text
            'the table structure of the output shapefile is the same as input
            pParameterArray.Add(filePath)
            pParameterArray.Add(name)
            pParameterArray.Add("Point")
            pParameterArray.Add(ComboBox1.Text)
            'if input has no coordinate szstem, also output won't have
            If Not pSpatialReference.FactoryCode = 0 Then
                pGP.SetEnvironmentValue("outputCoordinateSystem", pSpatialReference.FactoryCode)
            End If

            'execute the tool
            Dim pResult As IGeoProcessorResult = Nothing
            pResult = pGP.Execute("CreateFeatureClass_management", pParameterArray, Nothing)

            'prepare Cursor
            Dim pFeatureDB As IFeature = Nothing
            Dim pFCursorDB As IFeatureCursor = Nothing

            'read the new layer, created by the geoprocessing tool, and indexes of fields for dipdir and dip
            'it is the first layer in TOC, because it is a point geometry
            'search for the layer name is not a good idea, beacuse long names are shorten automatically
            Dim pFLTrendy As IFeatureLayer = Nothing
            pFLTrendy = pMap.Layer(0)

            Dim pFCTrendy As IFeatureClass = Nothing
            pFCTrendy = pFLTrendy.FeatureClass

            Dim pFeatureTrendy As IFeature = Nothing
            Dim pTableTrendy As ITable = Nothing
            Dim pFieldsTrendy As IFields = Nothing
            pTableTrendy = pFCTrendy
            pFieldsTrendy = pTableTrendy.Fields
            'find fields and read indexes
            Dim iAlfa, iFi As Integer
            iAlfa = pTableTrendy.FindField(ComboBox2.Text)
            iFi = pTableTrendy.FindField(ComboBox3.Text)
            If iAlfa < 0 Or iFi < 0 Then
                MsgBox("Some of the fields weren´t found in the input table.")
                Exit Sub
            End If

            'iALfa and iFi are the fields indexes, if the input is SDE data source, 
            'it has 2 fields less than normally (FID and SHAPE is missing)
            'Therefore input and output have different number of fields, which will be problem when writing into output table.
            'this difference can be solved by adding a constant "cor".
            'in case of sde data source sde = 2, otherwise cor = 0.
            'Dim cor As Long
            'cor = 0
            'If pFLDB.DataSourceType = "SDE Feature Class" Then cor = -2

            'variables for calculation of averaged orientation from orientation matrix
            Dim W, lAB, x, y, X2, Y2 As Double
            Dim m11, m12, m13, m22, m23, m33 As Double
            Dim Alfa, Fi, VeX, VeY, VeZ As Double
            Dim pPoint As IPoint = Nothing
            pPoint = New Point
            Dim aArray(1) As Double
            Dim bArray(2) As Double
            Dim sArray(5) As Double
            Dim pArray(2) As Double

            'set progressbar
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = 100

            'the next cycle creates net of averaging stations in the step of n (interval of av.stations) in the extent of selected input features
            For x = MinX To MaxX Step n

                'move progress bar
                ProgressBar1.Value = ((x - MinX) / (MaxX - MinX)) * 100

                For y = MinY To MaxY Step n

                    'null variables
                    m11 = 0
                    m12 = 0
                    m13 = 0
                    m22 = 0
                    m33 = 0
                    m23 = 0

                    'read selected features in the cycle and calculate orientation matrix from orientations 
                    'of all selected features
                    pFLayerDB.Search(Nothing, True, pFCursorDB)
                    pFeatureDB = pFCursorDB.NextFeature
                    Try
                        Do Until pFeatureDB Is Nothing
                            'lAB is distance of input feature from averaging station
                            'add 0,001 for case of more features with the same coordinates (to prevent lAB=0 -division by zero)
                            'X2, Y2 are coordinates of input feature
                            'lAB = 0
                            Try
                                X2 = pFeatureDB.Shape.Envelope.XMax
                                Y2 = pFeatureDB.Shape.Envelope.YMax
                                lAB = Sqrt(((X2 - x) ^ 2) + ((Y2 - y) ^ 2)) + 0.001

                                'orientation matrix is calculated only if input feature is within radius of influence.
                                If lAB < pv Then
                                    'W is the weigh parameter
                                    W = 1 / lAB
                                    'sumaW = sumaW + W
                                    'send dipdir and dip to function "Vektory_normala", that returns vector components of the normal
                                    aArray(0) = pFeatureDB.Value(iAlfa)
                                    aArray(1) = pFeatureDB.Value(iFi)
                                    bArray = Vektory_normala(aArray)
                                    VeX = bArray(0)
                                    VeY = bArray(1)
                                    VeZ = bArray(2)
                                    'calculate components of orientation matrix
                                    m11 = m11 + VeX * VeX * W
                                    m22 = m22 + VeY * VeY * W
                                    m33 = m33 + VeZ * VeZ * W
                                    m12 = m12 + VeX * VeY * W
                                    m13 = m13 + VeX * VeZ * W
                                    m23 = m23 + VeY * VeZ * W

                                End If
                            Catch ex As Exception

                            End Try
                            pFeatureDB = pFCursorDB.NextFeature
                        Loop
                    Catch ex As Exception
                        'Msgbox
                    End Try

                    'calculate averaged orientation and assign it to the averaging station
                    'if all features are outside the radius of influence, matrix is null
                    If Not m11 + m22 + m33 = 0 Then
                        'if matrix is not null, matrix components are read into sArray and sent it to Function
                        'Char_cisla_osa function read components of matrix and returns orientation of fold axis (trend/plunge)
                        sArray(0) = m11
                        sArray(1) = m22
                        sArray(2) = m33
                        sArray(3) = m12
                        sArray(4) = m13
                        sArray(5) = m23
                        'trend (Alfa) and plunge (Fi)
                        pArray = Char_cisla_osa(sArray)
                        Alfa = pArray(0)
                        Fi = pArray(1)

                        'create new feature in the new shapefile. XY and point geometry are assigned.
                        pFeatureTrendy = pFCTrendy.CreateFeature
                        pPoint.PutCoords(x, y)
                        pFeatureTrendy.Shape = pPoint
                        'values of averaged orientation are writen in the table
                        pFeatureTrendy.Value(iAlfa) = Alfa
                        pFeatureTrendy.Value(iFi) = Fi
                        pFeatureTrendy.Store()
                    End If
                Next y
            Next x


            'refresh
            'pMxDoc.ActiveView.Refresh()

            'set symbology of point features (averaged orientations) - oriented symbols
            Call AssignHollowRenderer3(pFLTrendy)

            'refresh
            pMxDoc.ActiveView.Refresh()

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Select workspace folder button
        Try
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
        'read from text file
        Try
            'Get Desktop Adress
            Dim filePath As String
            filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

            Dim fileContents As String
            'fileContents = My.Computer.FileSystem.ReadAllText("C:\temp\TableDef.txt")
            fileContents = My.Computer.FileSystem.ReadAllText(filePath & "\TableDef.txt")
            Dim b As String = Nothing
            Using MyReader As New  _
            Microsoft.VisualBasic.FileIO.TextFieldParser(filePath & "\TableDef.txt")

                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim CurrentRow As String()

                While Not MyReader.EndOfData

                    CurrentRow = MyReader.ReadFields()

                    ComboBox2.Text = CurrentRow(0)
                    ComboBox3.Text = CurrentRow(1)
                End While
            End Using



        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'write to text file
        'stF is field for control point id, tySF is type of S structure, dipdirSF and dipSF 
        'are Fields for dip direction and dip of S structures
        Try
            Dim dipdirSF, dipSF, tySF, stF As String
            dipdirSF = ComboBox2.Text
            dipSF = ComboBox3.Text
            tySF = " "
            stF = " "

            'Get Desktop adresss
            Dim filePAth As String
            filePAth = CreateObject("WScript.Shell").SpecialFolders("Desktop")

            My.Computer.FileSystem.WriteAllText(filePAth & "\TableDef.txt", (dipdirSF & "," & dipSF & "," & tySF & "," & stF), False)

        Catch ex As Exception
            Exit Sub

        End Try
    End Sub

    Private Sub FoldAxesForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'after the form is loaded, read the map and all layers from TOC into Combobox
            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            Dim aLName As String
            Dim I As Long
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap

            'read all layers from the active data frame
            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name
                ComboBox1.Items.Add(aLName)
            Next

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try

    End Sub


    Private Sub AssignHollowRenderer3(ByVal pGFLayer As IGeoFeatureLayer)
        Try
            'set symbology - oriented symbols
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document

            Dim pRender As ISimpleRenderer = pGFLayer.Renderer
            Dim pColor As IRgbColor = New RgbColor()
            pColor.Red = 0
            pColor.Green = 0
            pColor.Blue = 0

            Dim pMarkerSym As ICharacterMarkerSymbol = New CharacterMarkerSymbol()

            pMarkerSym.Size = 22
            pMarkerSym.Color = pColor
            pMarkerSym.CharacterIndex = 83

            'point symbol
            Dim pfont As New stdole.StdFont
            pfont.Name = "ESRI geology"
            pMarkerSym.Font = pfont

            'Dim aFont As IFontName = New FontName()
            'aFont.Name = "geo25_STR"
            'pMarkerSym.Font= aFont

            pRender.Symbol = pMarkerSym
            pGFLayer.Renderer = pRender

            'rotation
            Dim pRotRenderer As IRotationRenderer
            pRotRenderer = pGFLayer.Renderer

            pRotRenderer.RotationField = ComboBox2.Text

            pRotRenderer.RotationType = esriSymbolRotationType.esriRotateSymbolGeographic

            pMxDoc.UpdateContents()
            pMxDoc.ActiveView.Refresh()

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'open form
        Dim frm As New HelpForm
        frm.Show()
    End Sub
End Class