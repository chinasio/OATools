Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry
Imports System.Math
Imports ESRI.ArcGIS.DataSourcesRaster
Imports ESRI.ArcGIS.esriSystem
Imports System
Imports ESRI.ArcGIS.Geoprocessing

Public Class DensityForm

    Private Sub DensityForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'after form loads, get the map and read all layers
            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            Dim aLName As String
            Dim I As Long
            Dim pLayer As ILayer
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap

            'show all layers int the active DF in the combobox2
            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name
                ComboBox1.Items.Add(aLName)
                pLayer = pMap.Layer(I)
                If TypeOf pLayer Is IRasterLayer Then
                    ComboBox2.Items.Add(aLName)
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

 'after the input layer is selected
        'read the map and the input layer
        Try

            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap
            Dim pLDB As ILayer
            Dim pFLDB As IFeatureLayer

            Dim aLName As String
            Dim I As Long

            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next
            'get the input feature layer
            pLDB = pMap.Layer(I)
            pFLDB = pMap.Layer(I)

            'input layer must have data source
            If pFLDB.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            'input layer source is sde which is not allowed
            If pFLDB.DataSourceType = "SDE Feature Class" Then
                MsgBox("Input layer is a SDE Feature Class, which is not editable. Export it to shapefile or database Feature Class first.")
                Exit Sub
            End If

            'input layer mustn't be a raster layer 
            If TypeOf pLDB Is IRasterLayer Then
                MsgBox("You chose a raster layer.. but we need vector point feature class or shapefile!")
                Exit Sub
            End If

            'get the selection
            Dim pFClassDB As IFeatureClass
            pFClassDB = pFLDB.FeatureClass
            Dim pFSel As IFeatureSelection
            Dim pSelSet As ISelectionSet
            pFSel = pLDB
            pSelSet = pFSel.SelectionSet
            Label6.Text = pSelSet.Count & " features selected"
            
            If pSelSet.Count > 0 Then
                CheckBox1.Checked = True
                CheckBox1.Enabled = True
            End If

            If pSelSet.Count = 0 Then
                CheckBox1.Checked = False
                CheckBox1.Enabled = False
            End If

            'offer adequate smoothing parameter (depends on the number of input features)
            If pSelSet.Count > 0 Then
                If pSelSet.Count < 11 Then TextBox3.Text = 10
                If pSelSet.Count > 10 And pSelSet.Count < 31 Then TextBox3.Text = 20
                If pSelSet.Count > 30 And pSelSet.Count < 101 Then TextBox3.Text = 30
                If pSelSet.Count > 100 And pSelSet.Count < 251 Then TextBox3.Text = 50
                If pSelSet.Count > 250 Then TextBox3.Text = 65
            ElseIf pSelSet.Count = 0 Then
                If pFClassDB.FeatureCount(Nothing) < 11 Then TextBox3.Text = 10
                If pFClassDB.FeatureCount(Nothing) > 10 And pFClassDB.FeatureCount(Nothing) < 31 Then TextBox3.Text = 20
                If pFClassDB.FeatureCount(Nothing) > 30 And pFClassDB.FeatureCount(Nothing) < 101 Then TextBox3.Text = 30
                If pFClassDB.FeatureCount(Nothing) > 100 And pFClassDB.FeatureCount(Nothing) < 251 Then TextBox3.Text = 50
                If pFClassDB.FeatureCount(Nothing) > 250 Then TextBox3.Text = 65
            End If

            'get table and fields
            Dim pTableDB As ITable
            pTableDB = pFClassDB
            Dim pFieldsDB As IFields
            pFieldsDB = pTableDB.Fields

            'add the fields names to the combobox, add also option with no value.
            ComboBox3.Items.Add("")
            ComboBox4.Items.Add("")
            Dim j As Long
            For j = 0 To pFieldsDB.FieldCount - 1
                If pFieldsDB.Field(j).Type = esriFieldType.esriFieldTypeDouble Or pFieldsDB.Field(j).Type = esriFieldType.esriFieldTypeInteger Then
                    ComboBox3.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox4.Items.Add(pFieldsDB.Field(j).Name)
                End If
            Next

            Dim y, z As Long

            'fill the field names for dipdir and dip automatically if they have "smer_s" and "sklon_s" names
            y = pTableDB.FindField("DIPDIR")
            If y > (-1) Then
                ComboBox3.Text = "DIPDIR"
            End If
            z = pTableDB.FindField("DIP")
            If z > (-1) Then
                ComboBox4.Text = "DIP"
            End If

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'open form
        Dim frm As New HelpForm
        frm.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            'folder browser
            Dim FolderBrowserDialog1 As New FolderBrowserDialog
            With FolderBrowserDialog1

                '.RootFolder = Environment.SpecialFolder.MyDocuments
                .SelectedPath = "c:"
                .Description = "Select folder"
                'display selected path
                If .ShowDialog = DialogResult.OK Then
                    TextBox1.Text = (.SelectedPath)
                End If
            End With

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'read field names from text file, if already saved
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

                    ComboBox3.Text = CurrentRow(0)
                    ComboBox4.Text = CurrentRow(1)
                End While
            End Using

        Catch ex As Exception
            MsgBox("Exception. Did you set fields and remember setting first?")
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'write to text file
        'stF is field for control point id, tySF is type of S structure, dipdirSF and dipSF 
        'are Fields for dip direction and dip of S structures
        Try
            Dim dipdirSF, dipSF, tySF, stF As String
            dipdirSF = ComboBox3.Text
            dipSF = ComboBox4.Text
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

    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click
        'after switch to "Appearance" tab
        'read map and layers in TOC
        Try
            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            Dim aLName As String
            Dim I As Long
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap
            Dim pLayer As ILayer
            'read all layers into ComboBox
            For I = 0 To pMap.LayerCount - 1
                pLayer = pMap.Layer(I)
                If TypeOf pLayer Is IRasterLayer Then
                    'Dim pRL As IRasterLayer
                    'Set pRL = pLayer
                    'Dim pRasterBandCollection As IRasterBandCollection
                    'Set pRasterBandCollection = pRL.raster
                    'Dim pRasterBand As IRasterBand
                    'Set pRasterBand = pRasterBandCollection.Item(0)
                    'Dim pRasterStatistics As IRasterStatistics
                    'Set pRasterStatistics = pRasterBand.Statistics

                    'MsgBox pRasterStatistics.Maximum
                    aLName = pMap.Layer(I).Name
                    ComboBox2.Items.Add(aLName)
                End If
            Next
        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try
    End Sub



    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'start / calculate density distibution

        On Error GoTo eh

        'varning if some input boxes are not filled in the form
        If ComboBox1.Text = "" Or ComboBox4.Text = "" Or ComboBox3.Text = "" Or TextBox2.Text = "" Or TextBox1.Text = "" Then
            MsgBox("You didn´t choose some of the required parameters!")
            Exit Sub
        End If

        'read checkbox value (checked if is to be worked with selection, value = 1)
        Dim SelYes As Integer
        If CheckBox1.Checked = True Then
            SelYes = 1
        Else
            SelYes = 0
        End If

        'PART 1: MAP, INPUT LAYER, SELECTION, PARAMETERS, FIELDS, CURSOR
        'read map and the input layer
        Dim pMxDoc As IMxDocument
        pMxDoc = My.ArcMap.Document

        Dim pMap As IMap
        pMap = pMxDoc.FocusMap

        Dim pLayer As ILayer
        Dim aLName As String
        Dim io As Long

        For io = 0 To pMap.LayerCount - 1
            aLName = pMap.Layer(io).Name

            If aLName = ComboBox1.Text Then
                Exit For
            End If
        Next

        pLayer = pMap.Layer(io)

        'input layer mustn´t be a raster layer
        If TypeOf pLayer Is IRasterLayer Then
            MsgBox("You chose a raster layer..but we need a vector point feature class or shapefile!")
            Exit Sub
        End If

        'input layer must have data source
        If pLayer.Valid = False Then
            MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
            Exit Sub
        End If

        'constants for the case of calculation with fall lines (Checkbox2 = True)
        Dim SL_AZ, SL_IN As Double
        If CheckBox2.Checked = False Then
            SL_AZ = 0
            SL_IN = 0
        Else
            SL_AZ = 180
            SL_IN = 90
        End If

        'read the SelectionSet (from the input -tectonic- layer)
        Dim pFLayerA As IFeatureLayer
        pFLayerA = pMap.Layer(io)

        Dim pFSel As IFeatureSelection
        Dim pSelSet As ISelectionSet
        pFSel = pFLayerA
        'count number of selected features
        Dim n As Integer
        If SelYes = 1 Then
            pSelSet = pFSel.SelectionSet
            n = pSelSet.Count
        Else
            'in case of no selection is made, read all
            pFSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            pSelSet = pFSel.SelectionSet
            n = pSelSet.Count
            'cancel the selection
            pFSel.Clear()
        End If

        'find the fields for dip direction and dip in the input table
        Dim pFClassA As IFeatureClass
        pFClassA = pFLayerA.FeatureClass
        Dim pTableA As ITable
        pTableA = pFClassA
        'indexes for the fields
        Dim iAlfa, iFi As Integer
        iAlfa = pTableA.FindField(ComboBox3.Text)
        iFi = pTableA.FindField(ComboBox4.Text)

        'varn if input field weren´t found in the table
        If iAlfa < 0 Or iFi < 0 Then
            MsgBox("Some of the fields weren´t found in the input table.")
            Exit Sub
        End If

        'prepare Feature Cursor, feature cursor will go through all features 
        Dim pFCursorA As IFeatureCursor = Nothing
        Dim pFeatureA As IFeature = Nothing

        'PART2: CREATE RASTER WORKSPACE, RASTERDATASET, RASTER, IT´S PARAMETERS, PIXELBLOCK


        'create raster workspace
        Dim pWS As IRasterWorkspace2
        Dim pWSF As IWorkspaceFactory
        pWSF = New RasterWorkspaceFactory
        pWS = pWSF.OpenFromFile(TextBox1.Text, 0)

        'Define origin of the raster dataset
        Dim pOrigin As IPoint
        pOrigin = New Point
        pOrigin.PutCoords(-2.5, 0)

        'width and height of the raster
        Dim Width As Integer, Height As Integer
        Width = 50
        Height = 50
        'cell size in x and y direction
        Dim CellX As Double, CellY As Double
        CellX = 8
        CellY = 8


        'coordinate system
        'Dim sr As ISpatialReference
        'sr = New UnknownCoordinateSystem

        'create raster dataset using textbox2 for the file name
        Dim pDs As IRasterDataset2
        pDs = pWS.CreateRasterDataset(TextBox2.Text & ".tif", "TIFF", pOrigin, Width, Height, CellX, CellY, 1, rstPixelType.PT_DOUBLE, Nothing, True)

        'Get the raster band
        Dim rasterBandCollection As IRasterBandCollection
        Dim rasterProps As IRasterProps
        Dim band As IRasterBand
        rasterBandCollection = pDs
        band = rasterBandCollection.Item(0)
        rasterProps = band
        rasterProps.NoDataValue = 255


        'create a raster from the dataset
        Dim pRaster As IRaster
        pRaster = pDs.CreateFullRaster

        'create a pixel block using the weight and height of the raster dataset
        Dim pPB As IPixelBlock3
        Dim pPnt As IPnt 'blocksize
        pPnt = New Pnt
        pPnt.SetCoords(Width, Height)

        pPB = pRaster.CreatePixelBlock(pPnt)

        'PART 3: populate pixels with values to the pixel block

        Dim v As System.Array 'pixels
        v = CType(pPB.PixelData(0), System.Array)

        'variables: 
        Dim I As Integer, j As Integer  'pixel order in x and y direction
        'Dim k, d As Integer    '
        Dim r As Double  'radius
        Dim xo, yo As Integer  'coordinates of the diagram center point
        Dim Cn As Double  'smoothing parameter

        'read the smoothing parameter from TextBox2
        Cn = TextBox3.Text
        'xo,yo diagram center point
        xo = Round(Width / 2, 0)
        yo = Round(Height / 2, 0)
        'radius
        r = xo


        'variables needed for the density value calculation

        Dim Ro, xd, yd, x, y, h, Nu1, Nu2, Nu3, x3d, y3d, z3d, W, w1 As Double
        Const PI = 3.14159265
        Dim Fi, Alfa, VeX, VeY, VeZ As Double


        'null the matrix first
        For I = 1 To 2 * r - 1

            For j = 1 To 2 * r - 1
                'v(I, j) = 0
                v.SetValue(0, I, j)
            Next j

        Next I

        'set ProgressBar
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 100

        'arrays, sArray will contain dip dir and dip and sends it to a next function, 
        'that calculates vector components of the normal to plane
        Dim sArray(1) As Integer
        'tArray will contain results (vector components of the normal to plane)
        Dim tArray(2) As Double

        For I = 0 To 2 * r - 1
            ProgressBar1.Value = (I / Width) * 100
            'y direction of pixel (l,j) from the diagram center point
            yd = (I + 0.1 - yo) * Sqrt(2) / r

            For j = 0 To 2 * r - 1

                'null variables
                h = 0
                Nu1 = 0
                Nu2 = 0
                Nu3 = 0
                'n = 0
                Ro = 0
                x3d = 0
                y3d = 0
                z3d = 0
                W = 0

                'x direction of the pixel (l,j) from the diagram center
                'xd = Sqr(2) * (CellX * i - xo) / r
                xd = (j - xo) * Sqrt(2) / r + 0.01
                '0.01 is the correction (move left).
                xd = -xd
                'r is the radius of the sphere, diagram radius is Sqrt(2)
                'Ro is the direct distance of pixel (l, j) from the diagram center
                Ro = Sqrt((xd ^ 2) + (yd ^ 2))

                'when Ro> Sqrt(2) (=radius of the diagram), pixel (l, j) falls outside the diagram outline.
                If Ro < Sqrt(2) Then
                    Fi = PI / 2 - 2 * Atan(Ro / 2 / Sqrt((-Ro / 2) * (Ro / 2) + 1))
                    'Arcsin(x) = Atn(x / Sqr(-x * x + 1))

                    If xd = 0 Then
                        Alfa = Sign(yd) * PI / 2
                    Else
                        Alfa = (2 - Sign(yd) - Sign(xd * yd)) * (PI / 2) + Atan(yd / xd)

                    End If
                    'vector component of angluar direction of pixel (l, j)
                    x3d = Cos(Alfa) * Cos(Fi)
                    y3d = Sin(Alfa) * Cos(Fi)
                    z3d = Sin(Fi)

                    'clear cursor and feature
                    pFCursorA = Nothing
                    pFeatureA = Nothing

                    'set cursor at the first row of the input table and read the fisrt feature
                    pSelSet.Search(Nothing, True, pFCursorA)
                    pFeatureA = pFCursorA.NextFeature


                    'PART 3b: cyclus reads features in the table - density from all features is calculated
                    Do Until pFeatureA Is Nothing
                        'this matrix reads dip and dip direction
                        sArray(0) = pFeatureA.Value(iAlfa) + SL_AZ
                        sArray(1) = Abs(SL_IN - pFeatureA.Value(iFi))
                        'values are sent to function, that returns vector components of the normal to plane
                        tArray = Vektory_normala(sArray)
                        'components of vector (normal to plane)
                        VeX = tArray(0)
                        VeY = tArray(1)
                        VeZ = tArray(2)
                        'density on pixel (l, j ) is calculated
                        Nu1 = ((x3d * VeX) + (y3d * VeY) + (z3d * VeZ)) ^ 2
                        Nu2 = Nu1 * Cn
                        Nu3 = Exp(Nu2)
                        h = h + Nu3

                        'End If

                        pFeatureA = pFCursorA.NextFeature

                    Loop

                    'value of density is assigned to the pixel
                    v.SetValue(h, I, j)

                End If
            Next j

        Next I



        'PART  4: create pixel block, raster layer, raster renderer, add it to the map


        'pixels are assigned to pixel block
        pPB.PixelData(0) = CType(v, System.Array)
     
        'define the location of the upper left corner of the pixel block is to write
        Dim pPntOrigin As IPnt
        pPntOrigin = New Pnt
        pPntOrigin.SetCoords(0, 0)

        'write the pixel block
        Dim pEdit As IRasterEdit
        pEdit = pRaster
        pEdit.Write(pPntOrigin, pPB)


        'Dim pPnt2 As IPnt
        'pPnt2 = New Pnt
        'pPnt2.SetCoords(-2.5, 400)

        'Release rasterEdit explicitly.
        System.Runtime.InteropServices.Marshal.ReleaseComObject(pEdit)
        '_____________________________________________________

        'create raster layer
        Dim pLy As IRasterLayer = New RasterLayer

        '16.11.14 místo createfromraster dám create fromDataser
        'pLy.CreateFromRaster(pRaster)
        pLy.CreateFromDataset(pDs)

        'pMxDoc.FocusMap.AddLayer(pLy)
        pMxDoc.ActiveView.Refresh()
        'add layer to the map
        pMap.AddLayer(CType(pLy, ILayer))

        '_________________________________________________
        'calculate statistics
        ''parameters array
        'Dim pParameterArray As IVariantArray
        'pParameterArray = New VarArray
        'pParameterArray.Add(pDs)


        'call Geoprocessing tool
        'Dim pGP As IGeoProcessor
        'pGP = New GeoProcessor
        'pGP.OverwriteOutput = True

        'Dim pResult As IGeoProcessorResult
        'pResult = pGP.Execute("CalculateStatistics_management", pParameterArray, Nothing)
        '?__________________________________________

        'Release objects.
        v = Nothing
        pPB = Nothing
        pEdit = Nothing
        pRaster = Nothing
        pDs = Nothing
        pWS = Nothing
     
        'SET COLORRAMP in this PART

        ' Create renderer and QI RasterRenderer
        Dim pStretchRen As IRasterStretchColorRampRenderer
        pStretchRen = New RasterStretchColorRampRenderer
        Dim pRasRen As IRasterRenderer
        pRasRen = pStretchRen
        pRasRen.ResamplingType = rstResamplingTypes.RSP_BilinearInterpolation

        ' Set raster for the renderer and update
        pRasRen.Raster = pRaster
        pRasRen.Update()


        'type of rasterStretch
        Dim pRSType As IRasterStretch
        pRSType = pStretchRen
        pRSType.StretchType = esriRasterStretchTypesEnum.esriRasterStretch_MinimumMaximum
        pRasRen = pRSType

        ' Define two colors
        Dim pFromColor As IColor
        Dim pToColor As IColor
        pFromColor = New RgbColor
        pFromColor.RGB = RGB(255, 255, 255)
        pToColor = New RgbColor
        pToColor.RGB = RGB(230, 0, 0)

        ' Create color ramp
        Dim pRamp As IAlgorithmicColorRamp
        pRamp = New AlgorithmicColorRamp
        pRamp.Size = 255
        pRamp.FromColor = pFromColor
        pRamp.ToColor = pToColor
        pRamp.CreateRamp(True)

        ' Plug this colorramp into renderer and select a band
        pStretchRen.BandIndex = 0
        pStretchRen.ColorRamp = pRamp

        ' Update the renderer with new settings and plug into layer
        pRasRen.Update()
        pLy.Renderer = pStretchRen
        'refresh
        pMxDoc.ActiveView.Refresh()
        pMxDoc.UpdateContents()

        'Release memeory
        pMxDoc = Nothing
        'pRLayer = Nothing
        pRaster = Nothing
        pStretchRen = Nothing
        pRasRen = Nothing
        pRamp = Nothing
        pToColor = Nothing
        pFromColor = Nothing


        Exit Sub
eh:
        MsgBox(Err.Number & ": " & Err.Description)


    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try

            'black colorramp
            If ComboBox2.Text = "" Then
                MsgBox("Choose an input raster in the box above the button!")
                Exit Sub
            End If
            'read the map and the layer
            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap

            Dim aLName As String
            Dim j As Long

            For j = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(j).Name

                If aLName = ComboBox2.Text Then
                    Exit For
                End If
            Next

            'Get the input raster
            Dim pRLayer As IRasterLayer
            pRLayer = pMap.Layer(j)


            Dim pRaster As IRaster
            pRaster = pRLayer.Raster

            'Create classfy renderer and QI RasterRenderer interface
            Dim pClassRen As IRasterClassifyColorRampRenderer
            pClassRen = New RasterClassifyColorRampRenderer
            Dim pRasRen As IRasterRenderer
            pRasRen = pClassRen
            pRasRen.ResamplingType = rstResamplingTypes.RSP_BilinearInterpolation
            'Set raster for the render and update
            pRasRen.Raster = pRaster
            pClassRen.ClassCount = TextBox4.Text + 1
            pRasRen.Update()
            'colors
            Dim pRGB As New RgbColor
            pRGB.Red = 255
            pRGB.Green = 255
            pRGB.Blue = 255
            Dim pRGB2 As New RgbColor
            pRGB2.Red = 0
            pRGB2.Green = 0
            pRGB2.Blue = 0

            'Create a color ramp to use
            Dim pRamp As IAlgorithmicColorRamp
            pRamp = New AlgorithmicColorRamp
            pRamp.Size = TextBox4.Text + 1
            pRamp.FromColor = pRGB
            pRamp.ToColor = pRGB2
            pRamp.CreateRamp(True)

            'Create symbol for the classes
            Dim pFSymbol As IFillSymbol
            pFSymbol = New SimpleFillSymbol

            'pFSymbol.Color = pRGB
            'loop through the classes and apply the color and label
            Dim I As Integer
            For I = 0 To pClassRen.ClassCount - 1
                pFSymbol.Color = pRamp.Color(I)
                pClassRen.Symbol(I) = pFSymbol
                pClassRen.Label(I) = "Class" & CStr(I)
            Next I

            'Update the renderer and plug into layer
            pRasRen.Update()
            pRLayer.Renderer = pClassRen
            pMxDoc.ActiveView.Refresh()
            pMxDoc.UpdateContents()

            'Release memeory
            pMxDoc = Nothing
            pMap = Nothing
            pRLayer = Nothing
            pRaster = Nothing
            pRasRen = Nothing
            pClassRen = Nothing
            pRamp = Nothing
            pFSymbol = Nothing

        Catch ex As Exception
            MsgBox("Apply classified renderer for exported rasters only.")
        End Try

    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try

            'red classified renderer
            If ComboBox2.Text = "" Then
                MsgBox("Choose an input raster in the box above the button!")
                Exit Sub
            End If
            'get the map and input raster layer
            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap

            Dim aLName As String
            Dim j As Long

            For j = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(j).Name

                If aLName = ComboBox2.Text Then
                    Exit For
                End If
            Next

            'Get the input raster
            Dim pRLayer As IRasterLayer
            pRLayer = pMap.Layer(j)
            Dim pRaster As IRaster
            pRaster = pRLayer.Raster

            'Create classfy renderer and QI RasterRenderer interface
            Dim pClassRen As IRasterClassifyColorRampRenderer
            pClassRen = New RasterClassifyColorRampRenderer
            Dim pRasRen As IRasterRenderer
            pRasRen = pClassRen
            pRasRen.ResamplingType = rstResamplingTypes.RSP_BilinearInterpolation
            'Set raster for the render and update
            pRasRen.Raster = pRaster
            pClassRen.ClassCount = TextBox4.Text + 1
            pRasRen.Update()

            Dim pRGB As New RgbColor
            pRGB.Red = 255
            pRGB.Green = 255
            pRGB.Blue = 255

            Dim pRGB2 As New RgbColor
            pRGB2.Red = 168
            pRGB2.Green = 0
            pRGB2.Blue = 0

            'Create a color ramp to use
            Dim pRamp As IAlgorithmicColorRamp
            pRamp = New AlgorithmicColorRamp
            pRamp.Size = TextBox4.Text + 1
            pRamp.FromColor = pRGB
            pRamp.ToColor = pRGB2
            pRamp.CreateRamp(True)

            'Create symbol for the classes
            Dim pFSymbol As IFillSymbol
            pFSymbol = New SimpleFillSymbol
            
            'pFSymbol.Color = pRGB
            'loop through the classes and apply the color and label
            Dim I As Integer
            For I = 0 To pClassRen.ClassCount - 1
                pFSymbol.Color = pRamp.Color(I)
                pClassRen.Symbol(I) = pFSymbol
                pClassRen.Label(I) = "Class" & CStr(I)
            Next I

            'Update the renderer and plug into layer
            pRasRen.Update()
            pRLayer.Renderer = pClassRen
            pMxDoc.ActiveView.Refresh()
            pMxDoc.UpdateContents()

            'Release memeory
            pMxDoc = Nothing
            pMap = Nothing
            pRLayer = Nothing
            pRaster = Nothing
            pRasRen = Nothing
            pClassRen = Nothing
            pRamp = Nothing
            pFSymbol = Nothing


        Catch ex As Exception
            MsgBox("Apply classified renderer for exported rasters only.")
        End Try

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try

            'green classified renderer
            'get the map and input raster layer
            If ComboBox2.Text = "" Then
                MsgBox("Choose an input raster in the box above the button!")
                Exit Sub
            End If

            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap

            Dim aLName As String
            Dim j As Long

            For j = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(j).Name

                If aLName = ComboBox2.Text Then
                    Exit For
                End If
            Next


            'Get the input raster
            Dim pRLayer As IRasterLayer
            pRLayer = pMap.Layer(j)
            Dim pRaster As IRaster
            pRaster = pRLayer.Raster

            'Create classfy renderer and QI RasterRenderer interface
            Dim pClassRen As IRasterClassifyColorRampRenderer
            pClassRen = New RasterClassifyColorRampRenderer
            Dim pRasRen As IRasterRenderer
            pRasRen = pClassRen
            pRasRen.ResamplingType = rstResamplingTypes.RSP_BilinearInterpolation
            'Set raster for the render and update
            pRasRen.Raster = pRaster
            pClassRen.ClassCount = TextBox4.Text + 1
            pRasRen.Update()

            Dim pRGB As New RgbColor
            pRGB.Red = 255
            pRGB.Green = 255
            pRGB.Blue = 255

            Dim pRGB2 As New RgbColor
            pRGB2.Red = 38
            pRGB2.Green = 115
            pRGB2.Blue = 0

            'Create a color ramp to use
            Dim pRamp As IAlgorithmicColorRamp
            pRamp = New AlgorithmicColorRamp
            pRamp.Size = TextBox4.Text + 1
            pRamp.FromColor = pRGB
            pRamp.ToColor = pRGB2
            pRamp.CreateRamp(True)

            'Create symbol for the classes
            Dim pFSymbol As IFillSymbol
            pFSymbol = New SimpleFillSymbol

            'pFSymbol.Color = pRGB
            'loop through the classes and apply the color and label
            Dim I As Integer
            For I = 0 To pClassRen.ClassCount - 1
                pFSymbol.Color = pRamp.Color(I)
                pClassRen.Symbol(I) = pFSymbol
                pClassRen.Label(I) = "Class" & CStr(I)
            Next I

            'Update the renderer and plug into layer
            pRasRen.Update()
            pRLayer.Renderer = pClassRen
            pMxDoc.ActiveView.Refresh()
            pMxDoc.UpdateContents()

            'Release memeory
            pMxDoc = Nothing
            pMap = Nothing
            pRLayer = Nothing
            pRaster = Nothing
            pRasRen = Nothing
            pClassRen = Nothing
            pRamp = Nothing
            pFSymbol = Nothing


        Catch ex As Exception
            MsgBox("Apply classified renderer for exported rasters only.")
        End Try

    End Sub

  
End Class