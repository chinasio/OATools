Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geoprocessing
Imports System.Math

Public Class HistogramForm

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        'choose input layer
        Try
            'read index of selected layer
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap
            'Dim pLDB As ILayer
            Dim pFLayer As IFeatureLayer
            Dim aLName As String
            Dim I As Long

            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            pFLayer = pMap.Layer(I)
            Dim pFClass As IFeatureClass
            pFClass = pFLayer.FeatureClass

            Dim pTable As ITable
            pTable = pFClass
            Dim pFields As IFields
            pFields = pTable.Fields
            'Dim pField As IField

            'input layer must have data source
            If pFLayer.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If
            'input layer mustn't be of sde data source 
            If pFLayer.DataSourceType = "SDE Feature Class" Then
                MsgBox("Input layer is a SDE Feature Class, which is not editable. Export it to shapefile or database Feature Class first.")
                Exit Sub
            End If

            'input layer mustn't be a raster layer
            If TypeOf pMap.Layer(I) Is IRasterLayer Then
                MsgBox("You chose a raster layer.. but we need vector point feature class or shapefile!")
                Exit Sub
            End If
            'read selection
            Dim pFSelection As IFeatureSelection
            pFSelection = pFLayer

            Dim pSelSet As ISelectionSet
            pSelSet = pFSelection.SelectionSet

            If pSelSet.Count > 0 Then
                RadioButton1.Enabled = True
                RadioButton1.Checked = True
            End If

            If pSelSet.Count = 0 Then
                RadioButton1.Enabled = False
                RadioButton1.Checked = False
            End If

            If pFLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                ComboBox2.Enabled = True
                'add fields to combobox
                Dim j As Long
                For j = 1 To pFields.FieldCount - 1
                    If pFields.Field(j).Type = esriFieldType.esriFieldTypeDouble Or pFields.Field(j).Type = esriFieldType.esriFieldTypeInteger Then
                        ComboBox2.Items.Add(pFields.Field(j).Name)
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

    Private Sub HistogramForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'after the form loads, read the map
            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap
            'Dim pFeatureLayer As IFeatureLayer
            'raed layers from TOC into combobox
            Dim aLName As String
            Dim I As Long

            For I = 0 To pMap.LayerCount - 1

                If Not TypeOf pMap.Layer(I) Is IRasterLayer Then
                    'Set pFeatureLayer = pMap.Layer(I)
                    'If pFeatureLayer.FeatureClass.ShapeType = 3 Then
                    aLName = pMap.Layer(I).Name
                    ComboBox1.Items.Add(aLName)
                    'End If
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'conut histogram shape from input orientations in a point geometry shp or FC
        Try
            'combobox1,2 and textbox1 and 2 must be filled by the user
            If ComboBox1.Text = "" Or ComboBox2.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Then
                MsgBox("Select Input Feature Layer, workspace and type a name for the output shapefile!")
                Exit Sub
            End If
           

            'work with selection
            Dim SelYes As Integer
            SelYes = 0
            If RadioButton1.Checked = True Then
                SelYes = 1
            End If

            'read the input layer
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pActiveView As IActiveView
            pActiveView = pMxDoc.FocusMap
            Dim pMap As IMap
            'Set pMap = pMxDoc.FocusMap
            pMap = pActiveView
            Dim aLName As String
            Dim t As Long

            For t = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(t).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            'input layer must not be a raster layer
            If TypeOf pMap.Layer(t) Is IRasterLayer Then
                MsgBox("Input layer must not be raster!")
                Exit Sub
            End If

            Dim pFPointLayer As IFeatureLayer
            pFPointLayer = pMxDoc.FocusMap.Layer(t) 'read point geometry layer

            'input layer must have data source
            If pFPointLayer.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            'input layer must be point geometry
            If Not pFPointLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Or pFPointLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryMultipoint Then
                MsgBox("Input layer is not Point geometry!")
                Exit Sub
            End If

            'find field with dipdir
            Dim pPointClass As IFeatureClass
            pPointClass = pFPointLayer.FeatureClass
            Dim pPointFeat As IFeature
            Dim pTable As ITable
            Dim pFields As IFields
            pTable = pPointClass
            pFields = pTable.Fields
            Dim iAz As Integer
            iAz = pTable.FindField(ComboBox2.Text)
            If iAz = -1 Then
                MsgBox("Choose a field with dip direction in the box")
                Exit Sub
            End If

            'create a new shapefile using geoprocessing tool
            Dim filePath, name As String
            filePath = TextBox1.Text
            name = TextBox2.Text

            'call the tool
            Dim pSpatialReference As ISpatialReference
            Dim geoDataset As IGeoDataset
            geoDataset = pFPointLayer
            pSpatialReference = geoDataset.SpatialReference
            Dim pGP As IGeoProcessor
            pGP = New GeoProcessor
            pGP.OverwriteOutput = True

            'parameters array for "CreateFeatureClass"
            Dim pParameterArray As ESRI.ArcGIS.esriSystem.IVariantArray
            pParameterArray = New ESRI.ArcGIS.esriSystem.VarArray

            pParameterArray.Add(filePath)
            pParameterArray.Add(name)
            pParameterArray.Add("Polygon")
            pParameterArray.Add(ComboBox1.Text)

            'pokud nemá vstupní vrstva sdefinovaný souř. systém, nebude mít ani výsledná
            'If Not pSpatialReference.FactoryCode = 0 Then
            'pGP.SetEnvironmentValue("outputCoordinateSystem", pSpatialReference.FactoryCode)
            'End If
            'cooridnate system: Křovák (S-JTSK Krovak EastNorth), but it doesn´t matter what CS the histogram has.
            pGP.SetEnvironmentValue("outputCoordinateSystem", 102067)

            Dim pResult As IGeoProcessorResult
            pResult = pGP.Execute("CreateFeatureClass_management", pParameterArray, Nothing)

            Dim pFCursor As IFeatureCursor = Nothing
            Dim Fci As Double
            'if work with selection only, read the selction
            If SelYes = 1 Then
                Dim pFSelection As IFeatureSelection
                pFSelection = pFPointLayer
                Dim pInSelSet As ISelectionSet
                pInSelSet = pFSelection.SelectionSet
                pInSelSet.Search(Nothing, True, pFCursor)
                Fci = pInSelSet.Count
            Else
                pFCursor = pFPointLayer.Search(Nothing, True)
                Fci = pFPointLayer.FeatureClass.FeatureCount(Nothing)
            End If

            'read the first feature in the table of input layer
            pPointFeat = pFCursor.NextFeature
            'Dim I As Integer
            'definition of the diagram center point and the 10 degrees classes
            ''Dim dAzimuth As Double
            Dim ToPoint As IPoint
            Dim StartPoint As IPoint
            StartPoint = New Point
            StartPoint.X = 200
            StartPoint.Y = 200
            'Dim pLength As Double
            'Const PI = 3.14159265358979
            Dim PrumAz As Double
            PrumAz = 0
            Dim N0, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14, N15, N16, N17 As Integer
            N0 = 0
            N1 = 0
            N2 = 0
            N3 = 0
            N4 = 0
            N5 = 0
            N6 = 0
            N7 = 0
            N8 = 0
            N9 = 0
            N10 = 0
            N11 = 0
            N12 = 0
            N13 = 0
            N14 = 0
            N15 = 0
            N16 = 0
            N17 = 0
            Dim y As Integer
            y = 0
            Dim ci As Double
            ci = 0
            'move progress bar
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = 100

            'read azimuth from teh table
            Do Until pPointFeat Is Nothing
                'move progress bar
                ci = ci + 1
                ProgressBar1.Value = (ci / Fci) * 100
                'get strike from dipdir
                PrumAz = (pPointFeat.Value(iAz) + 90) Mod (180)
                'MsgBox PrumAz
                'locate the class the strike belongs to, remember
                Select Case PrumAz
                    Case 0 To 9
                        N0 = N0 + 1

                    Case 10 To 19
                        N1 = N1 + 1

                    Case 20 To 29
                        N2 = N2 + 1

                    Case 30 To 39
                        N3 = N3 + 1

                    Case 40 To 49
                        N4 = N4 + 1

                    Case 50 To 59
                        N5 = N5 + 1

                    Case 60 To 69
                        N6 = N6 + 1

                    Case 70 To 79
                        N7 = N7 + 1

                    Case 80 To 89
                        N8 = N8 + 1

                    Case 90 To 99
                        N9 = N9 + 1

                    Case 100 To 109
                        N10 = N10 + 1

                    Case 110 To 119
                        N11 = N11 + 1

                    Case 120 To 129
                        N12 = N12 + 1

                    Case 130 To 139
                        N13 = N13 + 1

                    Case 140 To 149
                        N14 = N14 + 1

                    Case 150 To 159
                        N15 = N15 + 1

                    Case 160 To 169
                        N16 = N16 + 1

                    Case 170 To 179
                        N17 = N17 + 1

                End Select

                pPointFeat = pFCursor.NextFeature
            Loop

            'MsgBox N1 & " " & N2 & " " & N3 & " " & N4 & " " & N5 & " " & N6 & " " & N7 & " " & N8 & " " & N9 & " " & N10 & " " & N11 & " " & N12 & " " & N13 & " " & N14 & " " & N15 & " " & N16 & " " & N17

            'find the class with the highest frequency
            Dim v(18) As Integer
            Dim MaxVal As Object
            Dim f As Long
            'v = (N0, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14, N15, N16, N17)
            v(0) = N0
            v(1) = N1
            v(2) = N2
            v(3) = N3
            v(4) = N4
            v(5) = N5
            v(6) = N6

            v(7) = N7
            v(8) = N8
            v(9) = N9
            v(10) = N10
            v(11) = N11
            v(12) = N12
            v(13) = N13

            v(14) = N14
            v(15) = N15
            v(16) = N16
            v(17) = N17


            MaxVal = 0

            'array "v" is frequency, matrix "u" are the angles of the center of each class and in matrix "d" are the relative lengths of classes
            For f = 0 To UBound(v)
                If v(f) > MaxVal Then MaxVal = v(f)
            Next f

            'MsgBox MaxVal

            If MaxVal = 0 Then
                MsgBox("No point was recognized")
                Exit Sub
            End If

            'v matici "v" jsou četnosti, v matici "u" úhly středu a v matici "d" rel. délky tříd
            'craete collection of points that will form the polygon.
            Dim pPointCollection As IPointCollection
            pPointCollection = New Polygon

            'pPointCollection.AddPoint StartPoint
            'find the polygon layer in tOC, that th polygon shape will be cretaed in. (it is not the first in TOC)
            Dim aLName2 As String
            Dim z As Long

            For z = 0 To pMap.LayerCount - 1
                aLName2 = pMap.Layer(z).Name

                If aLName2 = TextBox2.Text Then
                    Exit For
                End If
            Next

            Dim pFeatureLayer As IFeatureLayer
            pFeatureLayer = pMap.Layer(z)
            'Check the feature type of the layer
            If Not pFeatureLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Then Exit Sub
            Dim pFeatureP As IFeature

            Dim s As Long
            Dim uhel1 As Long
            Dim uhel2 As Long
            Dim d As Double
            Dim delka As Long
            Dim x As Integer

            'create points in proper interval in distance from center point that corresponds to the frequency.
            For s = 0 To 360 Step 10
                uhel1 = (s - 90) * (-1)
                uhel2 = uhel1 - 10
                'MsgBox uhel1
                'MsgBox uhel2
                x = (s Mod (180)) / 10
                'MsgBox x
                d = v(x)
                d = (200 * d / MaxVal)
                ' MsgBox d
                delka = Round(d, 0)
                'MsgBox "délka je: " & delka & " uhél je: " & uhel
                If Not delka = 0 Then
                    pPointCollection.AddPoint(StartPoint)

                    ToPoint = PtConstructAngleDistance(StartPoint, uhel1, delka)
                    pPointCollection.AddPoint(ToPoint)

                    ToPoint = PtConstructAngleDistance(StartPoint, uhel2, delka)
                    pPointCollection.AddPoint(ToPoint)

                    pPointCollection.AddPoint(StartPoint)
                    '5/2015 next section was put inside If condition, it was after it:
                    'create polygon shape from point collection
                    pFeatureP = pFeatureLayer.FeatureClass.CreateFeature
                    pFeatureP.Shape = pPointCollection
                    pFeatureP.Store()
                    Dim pPolygon As IPolygon
                    pPolygon = pFeatureP.Shape
                    pPolygon.Close()
                    '5/2105 next line added
                    pPointCollection.RemovePoints(0, 4)
                End If
               
            Next s
            pActiveView.Refresh()
        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'conut histogram shape from input orientations in a LINE geometry shp or FC
        Try

            'combobox1 anf textbox1 and 2 must be filled by the user
            If ComboBox1.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Then
                MsgBox("Select Input Feature Layer, workspace and type a name for the output shapefile!")
                Exit Sub
            End If

            Dim SelYes As Integer
            SelYes = 0
            If RadioButton1.Checked = True Then
                SelYes = 1
            End If


            'read input layer
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pActiveView As IActiveView
            pActiveView = pMxDoc.FocusMap
            Dim pMap As IMap
            'Set pMap = pMxDoc.FocusMap
            pMap = pActiveView
            Dim aLName As String
            Dim t As Long

            For t = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(t).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            'no raster layer
            If TypeOf pMap.Layer(t) Is IRasterLayer Then
                MsgBox("Input layer must not be raster!")
                Exit Sub
            End If

            Dim pFLineLayer As IFeatureLayer
            pFLineLayer = pMxDoc.FocusMap.Layer(t) 'vrstva se zlomy

            'no missing data source
            If pFLineLayer.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            'must be line geometry
            If Not pFLineLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
                MsgBox("Input layer is not polyline geometry!")
                Exit Sub
            End If

            'crete new shp
            Dim filePath, name As String
            filePath = TextBox1.Text
            name = TextBox2.Text

            'call Geoprocessing tool
            Dim pSpatialReference As ISpatialReference
            Dim geoDataset As IGeoDataset
            geoDataset = pFLineLayer
            pSpatialReference = geoDataset.SpatialReference
            Dim pGP As IGeoProcessor
            pGP = New GeoProcessor
            pGP.OverwriteOutput = True

            'paremeters array
            Dim pParameterArray As ESRI.ArcGIS.esriSystem.IVariantArray
            pParameterArray = New ESRI.ArcGIS.esriSystem.VarArray
            pParameterArray.Add(filePath)
            pParameterArray.Add(name)
            pParameterArray.Add("Polygon")
            pParameterArray.Add(ComboBox1.Text)

            'pokud nemá vstupní vrstva sdefinovaný souř. systém, nebude mít ani výsledná
            'If Not pSpatialReference.FactoryCode = 0 Then
            'pGP.SetEnvironmentValue("outputCoordinateSystem", pSpatialReference.FactoryCode)
            'End If
            'cooridnate system: Křovák (S-JTSK Krovak EastNorth), but it doesn´t matter what CS the histogram has.
            pGP.SetEnvironmentValue("outputCoordinateSystem", 102067)

            Dim pResult As IGeoProcessorResult
            pResult = pGP.Execute("CreateFeatureClass_management", pParameterArray, Nothing)


            Dim pLineFeat As IFeature
            Dim pFCursor As IFeatureCursor = Nothing
            Dim Fci As Double
            'read selection
            If SelYes = 1 Then
                Dim pFSelection As IFeatureSelection
                pFSelection = pFLineLayer
                Dim pInSelSet As ISelectionSet
                pInSelSet = pFSelection.SelectionSet
                pInSelSet.Search(Nothing, True, pFCursor)
                Fci = pInSelSet.Count
            Else
                pFCursor = pFLineLayer.Search(Nothing, True)
                Fci = pFLineLayer.FeatureClass.FeatureCount(Nothing)
            End If

            'read first feature
            pLineFeat = pFCursor.NextFeature

            Dim pPolyline As IPolyline
            Dim pSegColl As ISegmentCollection
            Dim pSubSegment As ISegment
            Dim I As Integer

            'diagram center and 10 degrees classes
            Dim dAzimuth As Double
            Dim ToPoint As IPoint
            Dim StartPoint As IPoint
            StartPoint = New Point
            StartPoint.X = 200
            StartPoint.Y = 200
            'Dim pLength As Double
            Const PI = 3.14159265358979
            Dim SumAZ, PrumAz As Double
            SumAZ = 0
            PrumAz = 0
            Dim N0, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14, N15, N16, N17 As Integer
            N0 = 0
            N1 = 0
            N2 = 0
            N3 = 0
            N4 = 0
            N5 = 0
            N6 = 0
            N7 = 0
            N8 = 0
            N9 = 0
            N10 = 0
            N11 = 0
            N12 = 0
            N13 = 0
            N14 = 0
            N15 = 0
            N16 = 0
            N17 = 0
            Dim y As Integer
            y = 0
            Dim ci As Double
            ci = 0
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = 100

            'calculate azimuth of line features
            Do Until pLineFeat Is Nothing
                'move progress bar
                ci = ci + 1
                ProgressBar1.Value = (ci / Fci) * 100

                'azimuth of line is averaged segments azimuth (PrumAZ) 
                pPolyline = pLineFeat.Shape
                pSegColl = New Polyline
                pSegColl.AddSegmentCollection(pPolyline)

                SumAZ = 0
                y = pSegColl.SegmentCount

                For I = 0 To (pSegColl.SegmentCount - 1)
                    pSubSegment = pSegColl.Segment(I)

                    dAzimuth = Get_Azimuth(pSubSegment)
                    dAzimuth = dAzimuth * 180 / PI
                    SumAZ = SumAZ + dAzimuth

                Next I
                'MsgBox "i je: " & i
                PrumAz = SumAZ / y
                PrumAz = Round(PrumAz, 0)
                PrumAz = PrumAz Mod (180)
                'MsgBox PrumAz

                'locate class the PrumAZ belongs to
                Select Case PrumAz
                    Case 0 To 10
                        N0 = N0 + 1

                    Case 10 To 19
                        N1 = N1 + 1

                    Case 20 To 29
                        N2 = N2 + 1

                    Case 30 To 39
                        N3 = N3 + 1

                    Case 40 To 49
                        N4 = N4 + 1

                    Case 50 To 59
                        N5 = N5 + 1

                    Case 60 To 69
                        N6 = N6 + 1

                    Case 70 To 79
                        N7 = N7 + 1

                    Case 80 To 89
                        N8 = N8 + 1

                    Case 90 To 99
                        N9 = N9 + 1

                    Case 100 To 109
                        N10 = N10 + 1

                    Case 110 To 119
                        N11 = N11 + 1

                    Case 120 To 129
                        N12 = N12 + 1

                    Case 130 To 139
                        N13 = N13 + 1

                    Case 140 To 149
                        N14 = N14 + 1

                    Case 150 To 159
                        N15 = N15 + 1

                    Case 160 To 169
                        N16 = N16 + 1

                    Case 170 To 179
                        N17 = N17 + 1

                End Select

                pLineFeat = pFCursor.NextFeature
            Loop
            'MsgBox N1 & " " & N2 & " " & N3 & " " & N4 & " " & N5 & " " & N6 & " " & N7 & " " & N8 & " " & N9 & " " & N10 & " " & N11 & " " & N12 & " " & N13 & " " & N14 & " " & N15 & " " & N16 & " " & N17
            'find class with the highest frequency
            Dim v(18) As Integer
            Dim MaxVal As Object
            Dim f As Long
            'array "v" is frequency, matrix "u" are the angles of the center of each class and in matrix "d" are the relative lengths of classes

            'v = Array(N0, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14, N15, N16, N17)
            v(0) = N0
            v(1) = N1
            v(2) = N2
            v(3) = N3
            v(4) = N4
            v(5) = N5
            v(6) = N6

            v(7) = N7
            v(8) = N8
            v(9) = N9
            v(10) = N10
            v(11) = N11
            v(12) = N12
            v(13) = N13

            v(14) = N14
            v(15) = N15
            v(16) = N16
            v(17) = N17

            MaxVal = 0

            For f = 0 To UBound(v)
                If v(f) > MaxVal Then MaxVal = v(f)
            Next f

            'MsgBox MaxVal
            If MaxVal = 0 Then
                MsgBox("No line was recognized as a fault")
                Exit Sub
            End If

            'array "v" is frequency, matrix "u" are the angles of the center of each class and in matrix "d" are the relative lengths of classes
            'v matici "v" jsou četnosti, v matici "u" úhly středu a v matici "d" rel. délky tříd

            'craete point collection that will form polygon shape
            Dim pPointCollection As IPointCollection
            pPointCollection = New Polygon

            'pPointCollection.AddPoint StartPoint
            'find the polygon layer in TOC.
            Dim aLName2 As String
            Dim z As Long

            For z = 0 To pMap.LayerCount - 1
                aLName2 = pMap.Layer(z).Name

                If aLName2 = TextBox2.Text Then
                    Exit For
                End If
            Next

            Dim pFeatureLayer As IFeatureLayer
            pFeatureLayer = pMap.Layer(z)
            'Check the feature type of the layer
            If Not pFeatureLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Then Exit Sub
            Dim pFeatureP As IFeature

            Dim s As Long
            Dim uhel1 As Long
            Dim uhel2 As Long
            Dim d As Double
            Dim delka As Long
            Dim x As Integer

            'create points in proper interval in distance from center point that corresponds to the frequency.
            'bude vytvářet body v intervalu velikosti tříd ve vzdálenosti od středu přímo úměrné četnosti.
            For s = 0 To 360 Step 10
                uhel1 = (s - 90) * (-1)
                uhel2 = uhel1 - 10
                'MsgBox uhel1
                'MsgBox uhel2
                x = (s Mod (180)) / 10
                'MsgBox x
                d = v(x)
                d = (200 * d / MaxVal)
                ' MsgBox d
                delka = Round(d, 0)

                'MsgBox "délka je: " & delka & " uhél je: " & uhel
                If Not delka = 0 Then
                    pPointCollection.AddPoint(StartPoint)

                    ToPoint = PtConstructAngleDistance(StartPoint, uhel1, delka)
                    pPointCollection.AddPoint(ToPoint)

                    ToPoint = PtConstructAngleDistance(StartPoint, uhel2, delka)
                    pPointCollection.AddPoint(ToPoint)

                    pPointCollection.AddPoint(StartPoint)
                    '5/2015 next section was put inside If condition, it was after it:
                    'create polygon from points
                    pFeatureP = pFeatureLayer.FeatureClass.CreateFeature
                    pFeatureP.Shape = pPointCollection
                    pFeatureP.Store()
                    Dim pPolygon As IPolygon
                    pPolygon = pFeatureP.Shape
                    pPolygon.Close()
                    'next line added, 5/2015:
                    pPointCollection.RemovePoints(0, 4)
                End If


            Next s

            pActiveView.Refresh()

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'show help form
        Dim frm As New HelpForm
        frm.Show()

    End Sub
End Class