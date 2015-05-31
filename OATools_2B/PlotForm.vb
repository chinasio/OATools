Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.IO.IsolatedStorage
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.esriSystem
Imports System.Math

Public Class PlotForm

    Private Sub PlotForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

        
        'read the map and layers from TOC
        Dim pMxDoc As IMxDocument
        pMxDoc = My.ArcMap.Document
        Dim pMap As IMap
        pMap = pMxDoc.FocusMap
        Dim aLName As String = Nothing
        Dim i As Long = Nothing
        'add layers to combobox
        For i = 0 To pMap.LayerCount - 1
            aLName = pMap.Layer(i).Name
            ComboBox1.Items.Add(aLName)
        Next
        Catch ex As Exception
            MsgBox("PlotForm_Load error. " & Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            'input layer is selected:
            'read the map and the layer

            If ComboBox1.Text = "" Then Exit Sub

            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap
            Dim pFLDB As IFeatureLayer = Nothing


            Dim aLName As String 'layer name
            Dim i As Long 'layer order in TOC
            For i = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(i).Name
                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next
            pFLDB = pMap.Layer(i)

            'input layer must not be a raster layer
            If TypeOf pMap.Layer(i) Is IRasterLayer Then
                MsgBox("You chose a raster layer.. but we need vector point feature class or shapefile!")
                Exit Sub
            End If
            Dim pLDB As ILayer
            pLDB = pFLDB

            'input layer must have data source
            If pLDB.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            'read selection
            Dim pFSel As IFeatureSelection = Nothing
            Dim pSelSet As ISelectionSet = Nothing
            pFSel = pLDB
            pSelSet = pFSel.SelectionSet

            'display number of selected features
            Label9.Text = pSelSet.Count


            If pSelSet.Count > 0 Then
                CheckBox1.Checked = True
                CheckBox1.Enabled = True
            End If

            If pSelSet.Count = 0 Then
                CheckBox1.Checked = False
                CheckBox1.Enabled = False
            End If

            'table, fields
            Dim pFClass As IFeatureClass = Nothing
            pFClass = pFLDB.FeatureClass
            Dim pTableDB As ITable = Nothing
            pTableDB = pFClass
            Dim pFieldsDB As IFields = Nothing
            pFieldsDB = pTableDB.Fields

            'add fields names into comboboxes
            ComboBox2.Items.Add("")
            ComboBox3.Items.Add("")
            ComboBox4.Items.Add("")
            ComboBox5.Items.Add("")
            ComboBox6.Items.Add("")
            ComboBox7.Items.Add("")
            ComboBox8.Items.Add("")

            Dim j As Long = Nothing
            For j = 0 To pFieldsDB.FieldCount - 1
                If Not pFieldsDB.Field(j).Type = esriFieldType.esriFieldTypeOID And Not pFieldsDB.Field(j).Type = esriFieldType.esriFieldTypeGeometry Then
                    ComboBox2.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox3.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox6.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox4.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox5.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox7.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox8.Items.Add(pFieldsDB.Field(j).Name)
                End If
            Next

            'automatically fill the text boxes with fields names if the table has recommended structure
            Dim W, x, y, z, L1, L2, L3 As Long
            W = pTableDB.FindField("point_no")
            x = pTableDB.FindField("s_type")
            y = pTableDB.FindField("dipdir")
            z = pTableDB.FindField("dip")
            L1 = pTableDB.FindField("l_type")
            L2 = pTableDB.FindField("trend")
            L3 = pTableDB.FindField("plunge")

            If W > (-1) Then
                ComboBox2.Text = "POINT_NO"
            End If
            If x > (-1) Then
                ComboBox3.Text = "S_TYPE"
            End If
            If y > (-1) Then
                ComboBox4.Text = "DIPDIR"
            End If
            If z > (-1) Then
                ComboBox5.Text = "DIP"
            End If
            If L1 > (-1) Then
                ComboBox6.Text = "L_TYPE"
            End If
            If L2 > (-1) Then
                ComboBox7.Text = "TREND"
            End If
            If L3 > (-1) Then
                ComboBox8.Text = "PLUNGE"
            End If
            Label8.Enabled = True
        Catch ex As Exception
            MsgBox("PlotForm, ComboBox1_SelectedIndexChanged error. " & Err.Number & ": " & Err.Description)
            Exit Sub
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'write to text file
        'stF is field for control point id, tySF is type of S structure, dipdirSF and dipSF 
        'are Fields for dip direction and dip of S structures
         Try
            Dim dipdirSF, dipSF, tySF, stF As String
            dipdirSF = ComboBox4.Text
            dipSF = ComboBox5.Text
            tySF = ComboBox3.Text
            stF = ComboBox2.Text

            'Get Desktop adresss
            Dim filePAth As String
            filePAth = CreateObject("WScript.Shell").SpecialFolders("Desktop")

            My.Computer.FileSystem.WriteAllText(filePAth & "\TableDef.txt", (dipdirSF & "," & dipSF & "," & tySF & "," & stF), False)

        Catch ex As Exception
            Exit Sub

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
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
                    ComboBox2.Text = CurrentRow(3)
                    ComboBox3.Text = CurrentRow(2)
                    ComboBox4.Text = CurrentRow(0)
                    ComboBox5.Text = CurrentRow(1)
                End While
            End Using
        Catch ex As Exception
            MsgBox("Exception. Did you set fields and remember setting first?")
        End Try
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'folder browser
        RadioButton2.Checked = True
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
            'Dim SaveFileDialog1 As New SaveFileDialog()
            'SaveFileDialog1.ShowDialog()
            'SaveFileDialog1.Filter = "shp files (*.shp)"
            'SaveFileDialog1.FilterIndex = 2
            'SaveFileDialog1.RestoreDirectory = True
            'SaveFileDialog1.InitialDirectory = "C:"
            'SaveFileDialog1.Title = "save shapefile"
            'SaveFileDialog1.CheckFileExists = True
            'SaveFileDialog1.CheckPathExists = True
            'SaveFileDialog1.DefaultExt = "shp"

            'If (SaveFileDialog1.ShowDialog() = DialogResult.OK) Then
            'TextBox1.Text = SaveFileDialog1.FileName
            'End If
        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        'when option "plot into an exisitng shp" is checked, layers fromTOC are read into combobox
        'read the map
        Try
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap
            Dim aLName As String = Nothing
            Dim i As Long = Nothing
            'read layers
            For i = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(i).Name
                ComboBox9.Items.Add(aLName)
            Next
        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'after the button "plot as points" is used
        Try
            'check if everything needed is filled
            If RadioButton3.Checked = True Then 'add to existing shp
                If ComboBox9.Text = "" Then
                    MsgBox("Choose an existing SHP!")
                    Exit Sub
                End If
            End If

            If RadioButton2.Checked = True And (TextBox1.Text = "" Or TextBox2.Text = "") Then
                MsgBox("Select the output workspace (folder) and type a name for the output shapefile!")
                Exit Sub
            End If

            If ComboBox1.Text = "" Then 'input point layer with tectonics must be filled
                MsgBox("Choose an input layer!")
                Exit Sub
            End If

            If ComboBox2.Text = "" Then 'field for control point ID must be filled
                MsgBox("Field for control point ID must be filled!")
                Exit Sub
            End If

            If Not ComboBox4.Text = "" And ComboBox3.Text = "" Then
                MsgBox("Select field for S type")
                Exit Sub
            End If
            If Not ComboBox7.Text = "" And ComboBox6.Text = "" Then
                MsgBox("Select field for L type")
                Exit Sub
            End If

            'read the map
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap
            'Dim pLDB As ILayer
            'Dim pFLDB As IFeatureLayer

            'read the layer
            Dim aLName As String
            Dim I As Long
            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name
                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            'input layer must not be a raster layer
            If TypeOf pMap.Layer(I) Is IRasterLayer Then
                MsgBox("You chose a raster layer.. but we need vector point feature class or shapefile!")
                Exit Sub
            End If

            Dim pFLayer As IFeatureLayer
            pFLayer = pMap.Layer(I)

            'input layer must have data source
            If pFLayer.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If
            'input layer must not be of a sde data source
            If pFLayer.DataSourceType = "SDE Feature Class" Then
                MsgBox("Input layer is a SDE Feature Class. Export it to shapefile or database feature class first")
                Exit Sub
            End If

            'table, fields
            Dim pFClass As IFeatureClass = Nothing
            pFClass = pFLayer.FeatureClass

            Dim pTable As ITable = Nothing
            pTable = pFClass

            Dim pFields As IFields = Nothing
            pFields = pTable.Fields

            '9.4.2015: N field is added to the new shapefile, not the original table
            'add "N" field
            'Dim islN As Integer = Nothing
            'islN = pFields.FindField("N")
            'If islN = -1 Then
            'Dim pFieldEditN As IFieldEdit
            'pFieldEditN = New Field
            'pFieldEditN.Name_2 = "N"
            'pFieldEditN.Type_2 = esriFieldType.esriFieldTypeInteger
            'pFClass.AddField(pFieldEditN)
            'End If

            Dim SelYes As Integer = Nothing
            'read options
            If CheckBox1.Checked = True Then 'Plot only selected features
                SelYes = 1 'work with selection
            Else
                SelYes = 0 'work with all
            End If

            Dim shpYes As Integer = Nothing
            If RadioButton2.Checked = True Then 'create new shp option
                shpYes = 1
            End If

            If RadioButton3.Checked = True Then 'plot in existing shp option
                shpYes = 2
            End If

            'read selection
            Dim pFSelection As IFeatureSelection = Nothing
            pFSelection = pFLayer
            Dim pInSelSet As ISelectionSet = Nothing
            pInSelSet = pFSelection.SelectionSet

            Dim pPoint As IPoint = Nothing

            Dim cbod, smers, sklons, styp, smerl, sklonl, ltyp, filePath, name, Par4 As String
            cbod = ComboBox2.Text 'control point ID field
            smers = ComboBox4.Text 'dip direction field
            sklons = ComboBox5.Text 'dip field
            smerl = ComboBox7.Text 'trend
            sklonl = ComboBox8.Text 'plunge
            styp = ComboBox3.Text 'type of S structure
            ltyp = ComboBox6.Text 'type of L structure
            filePath = TextBox1.Text 'new shp workspace
            name = TextBox2.Text 'new shp name
            Par4 = ComboBox1.Text 'input layer name

            Dim VeX_s, VeY_s, VeZ_s, VeX_l, VeY_l, VeZ_l As Double
            Const r = 200 'radius of the projection
            Const xo = 200 'projection center point x coordinate
            Const yo = 200 'projection center point y coordinate

            Dim Ro, xd, yd, x, y As Double
            Dim ieN, iAZ_S, iSk_S, iAZ_L, iSk_L, iStyp, iLtyp, iDB As Integer
            ieN = 0
            'field indexes
            iAZ_S = 0 'dipdir
            iSk_S = 0 'dip
            iAZ_L = 0 'trend
            iSk_L = 0 'plunge
            iStyp = 0 'type of S structure
            iLtyp = 0 'type of L structure
            iDB = 0 'control point ID
            'find fields, read their indexes, they are -1, if they are not filled

            Dim pFeature As IFeature = Nothing
            iDB = pFields.FindField(cbod)
            If Not iDB > -1 Then
                MsgBox("Field " & cbod & " wasn´t found int the input table")
                Exit Sub
            End If

            'indexes of the fields
            If Not smers = "" Then
                iAZ_S = pFields.FindField(smers)
                iSk_S = pFields.FindField(sklons)
                iStyp = pFields.FindField(styp)
                If iAZ_S = -1 Or iSk_S = -1 Or iStyp = -1 Then
                    MsgBox("Some of the fields for S structure weren´t found")
                    Exit Sub
                End If
            End If

            If Not smerl = "" Then
                iAZ_L = pFields.FindField(smerl)
                iSk_L = pFields.FindField(sklonl)
                iLtyp = pFields.FindField(ltyp)
                If iAZ_L = -1 Or iSk_L = -1 Or iLtyp = -1 Then
                    MsgBox("Some of the fields for L structure weren´t found")
                    Exit Sub
                End If
            End If

            If iAZ_S = -1 Then
                Exit Sub
            End If
            If iAZ_L = -1 Then
                Exit Sub
            End If


            'iAZ_S = pFields.FindField(smers)
            'iSk_S = pFields.FindField(sklons)
            'iStyp = pFields.FindField(styp)
            'if wrong text string is set, exit
            'If iAZ_S = -1 Or iSk_S = -1 Then
            'MsgBox("Field for dipdir or dip of a S structure wasn´t found")
            'Exit Sub
            'End If
            'fields for L type are also optional, they are -1 if not filled
            'iAZ_L = pFields.FindField(smerl)
            'iSk_L = pFields.FindField(sklonl)
            'iLtyp = pFields.FindField(ltyp)

            'Create cursor and read the first feature
            Dim pFCursor As IFeatureCursor = Nothing
            'Set pFCursor = pFClass.Update(Nothing, True)
            'předchozí řádek plus UpdateFeature místo pFeature.Store použiju, pokud chci vynášet všechny měření, ne jen selection

            If SelYes = 1 Then
                pInSelSet.Search(Nothing, True, pFCursor)
            Else
                pFCursor = pFClass.Update(Nothing, True)
            End If
            pFeature = pFCursor.NextFeature

            Dim pElement As IElement = Nothing
            Dim pGraphicsContainer As IGraphicsContainer = Nothing
            Dim pSymbol As IMarkerSymbol = Nothing
            Dim pMElement As IMarkerElement = Nothing
            Dim pNewFLayer As IFeatureLayer = Nothing
            Dim pNewFClass As IFeatureClass = Nothing
            Dim pNewFeature As IFeature = Nothing
            Dim pNewTable As ITable = Nothing
            Dim pNewFields As IFields = Nothing
            Dim pFieldEditN As IFieldEdit = Nothing

            Dim sArray(1) As Integer
            Dim lArray(1) As Integer
            Dim kArray(2) As Double
            Dim mArray(2) As Double

            'If IsNull(pFeature.Value(iStyp)) = True Then
            If shpYes = 1 Then 'create new shp
                'Dim pSpatialReference As ISpatialReference
                'Dim geoDataset As IGeoDataset
                'geoDataset = pFLayer
                'pSpatialReference = geoDataset.SpatialReference

                'call Geoprocessing tool
                Dim pGP As IGeoProcessor
                pGP = New GeoProcessor
                pGP.OverwriteOutput = True

                'parameters array for "CreateFeatureClass" tool
                Dim pParameterArray As IVariantArray
                pParameterArray = New VarArray

                pParameterArray.Add(filePath)
                pParameterArray.Add(name)
                pParameterArray.Add("Point")
                pParameterArray.Add(Par4)

                'pokud nemá vstupní vrstva definovaný souř. systém, nebude mít ani výsledná
                'If Not pSpatialReference.FactoryCode = 0 Then
                'tady se přiděloval souř systém vstupní vrstvy novému shp
                'pGP.SetEnvironmentValue("outputCoordinateSystem", pSpatialReference.FactoryCode)
                '''''End If
                '23.1.2013 pevně dávám novému shp Křováka (kód 10267), protože, kdyby input wgs, output by byl taky a ten
                ' má počátek souř. syst. jinde než Křovák a vynesené body by nespadly do kruhu diagramu.

                Try
                    'coordinate system for plotted points is set
                    '7.12.2014 ruším souř.systém, výsledek nemá ss. Křovak byl matoucí.
                    'pGP.SetEnvironmentValue("outputCoordinateSystem", 102067)
                    'vytvoř novou třídu prvků
                    Dim pResult As IGeoProcessorResult
                    pResult = pGP.Execute("CreateFeatureClass_management", pParameterArray, Nothing)

                Catch ex As Exception
                    MsgBox("error creating new shapefile. " & Err.Number & ": " & Err.Description)
                    Exit Sub
                End Try

                'read the new shp
                pNewFLayer = pMap.Layer(0)
                pNewFClass = pNewFLayer.FeatureClass
                pNewTable = pNewFClass
                pNewFields = pNewTable.Fields


                'add "N" field
                Dim islN As Integer = Nothing
                islN = pFields.FindField("N")
                If islN = -1 Then
                    pFieldEditN = New Field
                    pFieldEditN.Name_2 = "N"
                    pFieldEditN.Type_2 = esriFieldType.esriFieldTypeInteger
                    pNewFClass.AddField(pFieldEditN)

                End If

            ElseIf shpYes = 2 Then
                'if "add to an existing shp" is selected, read the selected layer pNewFLayer 
                Dim l As Long
                For l = 0 To pMap.LayerCount - 1
                    aLName = pMap.Layer(l).Name
                    If aLName = ComboBox9.Text Then
                        Exit For
                    End If
                Next

                pNewFLayer = pMap.Layer(l)
                pNewFClass = pNewFLayer.FeatureClass
                pNewTable = pNewFClass
                pNewFields = pNewTable.Fields
                'layer must be point geometry
                If Not pNewFLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                    MsgBox("You haven´t chosen point layer to add points in. Choose point shapefile!")
                    Exit Sub
                End If
                'table structure must be the same as original
                If Not styp = "" Then
                    If pNewFClass.Fields.FindField(smers) = -1 Or pNewFClass.Fields.FindField(sklons) = -1 Or pNewFClass.Fields.FindField("N") = -1 Or pNewFClass.Fields.FindField(cbod) = -1 Or pNewFClass.Fields.FindField(styp) = -1 Then
                        MsgBox("Sorry, you chose shp with table structure that differs from input layer. Or integer N field is missing.")
                        Exit Sub
                    End If
                End If
                If Not ltyp = "" Then
                    If pNewFClass.Fields.FindField(smerl) = -1 Or pNewFClass.Fields.FindField(sklonl) = -1 Or pNewFClass.Fields.FindField("N") = -1 Or pNewFClass.Fields.FindField(cbod) = -1 Or pNewFClass.Fields.FindField(ltyp) = -1 Then
                        MsgBox("Sorry, you chose shp with table structure that differs from input layer. Or integer N field is missing.")
                        Exit Sub
                    End If
                End If
            End If
            'calculate xy coordinates in diagram DF and plot point, set symbology
            Do Until pFeature Is Nothing
                'planar tectonic orientations
                If Not iStyp = 0 Then

                    'calculate vector from angular orientation
                    sArray(0) = pFeature.Value(iAZ_S)
                    sArray(1) = pFeature.Value(iSk_S)

                    kArray = Vektory_normala(sArray)
                    VeX_s = kArray(0)
                    VeY_s = kArray(1)
                    VeZ_s = kArray(2)

                    'další výpočty na každém měření vypočte ro, xd a yd, přičte k xo a yo
                    ' vytvořím point geometry a vynesem

                    'calculate coordinates, create point
                    Ro = r * (1 - VeZ_s) ^ (1 / 2)
                    yd = r * VeX_s / (1 + VeZ_s) ^ (1 / 2)
                    xd = r * VeY_s / (1 + VeZ_s) ^ (1 / 2)
                    x = xo + xd
                    y = yo + yd
                    pPoint = New Point
                    pPoint.X = x
                    pPoint.Y = y
                    'create new feature if required and fill attributes
                    If shpYes = 1 Or shpYes = 2 Then

                        pNewFeature = pNewFClass.CreateFeature
                        pNewFeature.Shape = pPoint

                        pNewFeature.Value(iAZ_S) = pFeature.Value(iAZ_S)
                        pNewFeature.Value(iSk_S) = pFeature.Value(iSk_S)
                        pNewFeature.Value(iStyp) = pFeature.Value(iStyp)
                        pNewFeature.Value(iDB) = pFeature.Value(iDB)

                        pNewFeature.Value(pNewFields.FindField("N")) = pFeature.Value(0)
                        pNewFeature.Store()

                    Else
                        'plot as graphics if required
                        pElement = New MarkerElement
                        pElement.Geometry = pPoint
                        pGraphicsContainer = pMxDoc.FocusMap
                        pGraphicsContainer.AddElement(pElement, 0)

                        'symbol
                        pSymbol = New SimpleMarkerSymbol
                        Dim pModra2 As IRgbColor
                        pModra2 = New RgbColor
                        pModra2.Red = 0
                        pModra2.Green = 92
                        pModra2.Blue = 230
                        pSymbol.Color = pModra2
                        pSymbol.Size = 2
                        pMElement = pElement
                        pMElement.Symbol = pSymbol
                    End If
                End If
                'linear tectonic orientations:
                If Not iLtyp = 0 Then
                    If pFeature.Value(iAZ_L) <> 0 Then
                        'calculate vector from angular orientation
                        lArray(0) = pFeature.Value(iAZ_L)
                        lArray(1) = pFeature.Value(iSk_L)
                        mArray = Vektory_lineace(lArray)
                        VeX_l = mArray(0)
                        VeY_l = mArray(1)
                        VeZ_l = mArray(2)
                        'calculate coordinates, create point
                        Ro = r * (1 - VeZ_l) ^ (1 / 2)
                        yd = r * VeX_l / (1 + VeZ_l) ^ (1 / 2)
                        xd = r * VeY_l / (1 + VeZ_l) ^ (1 / 2)
                        x = xo + xd
                        y = yo + yd
                        pPoint = New Point
                        pPoint.X = x
                        pPoint.Y = y

                        If shpYes = 1 Or shpYes = 2 Then
                            'create new feature if required and fill attributes
                            pNewFeature = pNewFClass.CreateFeature
                            pNewFeature.Shape = pPoint
                            pNewFeature.Value(iAZ_L) = pFeature.Value(iAZ_L)
                            pNewFeature.Value(iSk_L) = pFeature.Value(iSk_L)
                            pNewFeature.Value(iLtyp) = pFeature.Value(iLtyp)
                            pNewFeature.Value(iDB) = pFeature.Value(iDB)
                            pNewFeature.Value(pNewFields.FindField("N")) = pFeature.Value(0)
                            pNewFeature.Store()

                        Else
                            'plot as graphics if required
                            pElement = New MarkerElement
                            pElement.Geometry = pPoint
                            pGraphicsContainer = pMxDoc.FocusMap
                            pGraphicsContainer.AddElement(pElement, 0)
                            'symbol
                            pSymbol = New SimpleMarkerSymbol

                            Dim pCervena As IRgbColor
                            pCervena = New RgbColor
                            pCervena.Red = 255
                            pCervena.Green = 0
                            pCervena.Blue = 0

                            pSymbol.Color = pCervena
                            pSymbol.Size = 2

                            pMElement = pElement
                            pMElement.Symbol = pSymbol
                        End If
                    End If
                End If
                pFeature = pFCursor.NextFeature
            Loop
            'refresh
            Dim pActiveView As IActiveView
            pActiveView = pMxDoc.FocusMap
            pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'fold axis calculation
        'výpočet osy vrásy z vybraných prvků bez závislosti na lokalizaci
        'check inputs
        Try
            If ComboBox1.Text = "" Then
                MsgBox("Choose an input layer with tectonics!")
                Exit Sub
            End If
            If ComboBox2.Text = "" Then
                MsgBox("Choose a field with control point ID!")
                Exit Sub
            End If

            'read map, input layer
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap = Nothing
            pMap = pMxDoc.FocusMap
            Dim pLDB As ILayer = Nothing
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

            'načtení proměnných pro názvy sloupců tabulky
            'Dim smers As String
            'Dim sklons As String
            'smers = ComboBox4.Value
            'sklons = ComboBox5.Value

            'read selection
            Dim pFSelection As IFeatureSelection = Nothing
            pFSelection = pFLDB
            Dim pFLayerDB As ISelectionSet = Nothing
            pFLayerDB = pFSelection.SelectionSet
            If pFLayerDB.Count = 1 Then
                MsgBox("Select two features minimally!")
                Exit Sub
            End If
            'if no selection, select all
            If pFLayerDB.Count = 0 Then
                pFSelection.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                pFLayerDB = pFSelection.SelectionSet
                'clear selection
                pFSelection.Clear()
            End If

            'read table,fields, fields names
            Dim pFClassDB As IFeatureClass = Nothing
            pFClassDB = pFLDB.FeatureClass
            Dim pTableDB As ITable = Nothing
            pTableDB = pFClassDB
            Dim pFeatureDB As IFeature = Nothing
            Dim pFCursorDB As IFeatureCursor = Nothing
            If ComboBox4.Text = "" Or ComboBox5.Text = "" Then
                MsgBox("Choose fields for dip and dip direction of the S structure!")
                Exit Sub
            End If

            'variables
            Dim m11, m12, m13, m22, m23, m33 As Double
            Dim Alfa, Fi, VeX, VeY, VeZ As Double
            'Const PI = 3.141592654
            Dim aArray(1) As Integer
            Dim bArray(2) As Double
            'null
            m11 = 0
            m12 = 0
            m13 = 0
            m22 = 0
            m33 = 0
            m23 = 0
            Fi = 0
            Alfa = 0
            'check input field names
            Dim iAlfa, iFi As Integer
            If Not ComboBox3.Text = "" Then
                iAlfa = pTableDB.FindField(ComboBox4.Text)
                iFi = pTableDB.FindField(ComboBox5.Text)
                If iAlfa = -1 Or iFi = -1 Then
                    MsgBox("Some of the fields for S structure weren´t found in the input table.")
                    Exit Sub
                End If
            ElseIf Not ComboBox6.Text = "" Then
                iAlfa = pTableDB.FindField(ComboBox7.Text)
                iFi = pTableDB.FindField(ComboBox8.Text)
                If iAlfa = -1 Or iFi = -1 Then
                    MsgBox("Some of the fields for L structure weren´t found in the input table.")
                    Exit Sub
                End If
            Else
                MsgBox("Choose the fields containing type of structure, dip and dip direction!")
                Exit Sub
            End If

            'orientation matrix is calculated from orientations in the table
            'read the first feature
            pFLayerDB.Search(Nothing, True, pFCursorDB)
            pFeatureDB = pFCursorDB.NextFeature

            Do Until pFeatureDB Is Nothing
                'calculate vector components from angular orientation
                aArray(0) = pFeatureDB.Value(pFeatureDB.Fields.FindField(ComboBox4.Text))
                aArray(1) = pFeatureDB.Value(pFeatureDB.Fields.FindField(ComboBox5.Text))
                bArray = Vektory_normala(aArray)
                VeX = bArray(0)
                VeY = bArray(1)
                VeZ = bArray(2)
                'components of orientation matrix
                m11 = m11 + VeX * VeX
                m22 = m22 + VeY * VeY
                m33 = m33 + VeZ * VeZ
                m12 = m12 + VeX * VeY
                m13 = m13 + VeX * VeZ
                m23 = m23 + VeY * VeZ

                pFeatureDB = pFCursorDB.NextFeature
            Loop

            'na uhlopříčce je součet složek matice rovný počtu měření při nevážené matici, u vážené součtu vah
            'do funkce pošle matici sArray, funkce vrátí matici o jednom řádku a dvou sloupcích (Alfa, Fi)
            Dim sArray(5) As Double
            sArray(0) = m11
            sArray(1) = m22
            sArray(2) = m33
            sArray(3) = m12
            sArray(4) = m13
            sArray(5) = m23
            'calculate eigenvector
            Dim pArray(1) As Double
            pArray = Char_cisla_osa(sArray)

            Alfa = pArray(0)
            Fi = pArray(1)
            'display orientation in angles in the form
            TextBox4.Text = Alfa
            TextBox5.Text = Fi

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            '"Calculate averaged orientation" button.
            'check inputs
            If ComboBox1.Text = "" Then
                MsgBox("Choose an input layer with tectonics!")
                Exit Sub
            End If
            If ComboBox2.Text = "" Then
                MsgBox("Choose a field with control point ID!")
                Exit Sub
            End If

            'read map,layers
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
            pFLDB = pMap.Layer(I)
            'read selection
            Dim pFSelection As IFeatureSelection
            pFSelection = pFLDB

            Dim pFLayerDB As ISelectionSet
            pFLayerDB = pFSelection.SelectionSet

            If pFLayerDB.Count = 0 Then
                pFSelection.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
                pFLayerDB = pFSelection.SelectionSet
                'clear selection
                pFSelection.Clear()
            End If

            'read table, fields
            Dim pFClassDB As IFeatureClass
            pFClassDB = pFLDB.FeatureClass
            Dim pTableDB As ITable
            pTableDB = pFClassDB
            Dim pFeatureDB As IFeature
            Dim pFCursorDB As IFeatureCursor = Nothing

            Dim iAlfa, iFi As Integer
            'check input fields
            If Not ComboBox3.Text = "" Then
                iAlfa = pTableDB.FindField(ComboBox4.Text)
                iFi = pTableDB.FindField(ComboBox5.Text)
                If iAlfa = -1 Or iFi = -1 Then
                    MsgBox("Some of the fields for S structure weren´t found in the input table.")
                    Exit Sub
                End If
            ElseIf Not ComboBox6.Text = "" Then
                iAlfa = pTableDB.FindField(ComboBox7.Text)
                iFi = pTableDB.FindField(ComboBox8.Text)
                If iAlfa = -1 Or iFi = -1 Then
                    MsgBox("Some of the fields for L structure weren´t found in the input table.")
                    Exit Sub
                End If
            Else
                MsgBox("Choose the names of fields containing type of structure, dip and dip direction!")
                Exit Sub
            End If

            'deklarace proměnných pro výpočet průměrného směru z matice orientace
            'variables
            Dim m11, m12, m13, m22, m23, m33 As Double
            Dim Alfa, Fi, VeX, VeY, VeZ As Double
            Dim aArray(1) As Integer
            Dim bArray(2) As Double
            'null
            m11 = 0
            m12 = 0
            m13 = 0
            m22 = 0
            m33 = 0
            m23 = 0
            Fi = 0
            Alfa = 0

            'read the first feature
            pFLayerDB.Search(Nothing, True, pFCursorDB)
            pFeatureDB = pFCursorDB.NextFeature

            'craete orientation matrix from orientations in the table
            Do Until pFeatureDB Is Nothing
                aArray(0) = pFeatureDB.Value(iAlfa)
                aArray(1) = pFeatureDB.Value(iFi)
                bArray = Vektory_normala(aArray)
                VeX = bArray(0)
                VeY = bArray(1)
                VeZ = bArray(2)
                m11 = m11 + VeX * VeX
                m22 = m22 + VeY * VeY
                m33 = m33 + VeZ * VeZ
                m12 = m12 + VeX * VeY
                m13 = m13 + VeX * VeZ
                m23 = m23 + VeY * VeZ
                pFeatureDB = pFCursorDB.NextFeature
            Loop
            'součet složek matice na uhlopříčce se musí rovnat počtu měření (u nevážené matice)
            'MsgBox(m11 + m22 + m33)
            'do funkce pošle matici sArray, funkce vrátí matici o jednom řádku a dvou sloupcích (Alfa, Fi)
            Dim sArray(5) As Double 'matrix components
            Dim pArray(1) As Double 'averaged orientation dipdir/dip
            Dim gArray(4) As Double 'eigenvalues, shape and strength parameter

            sArray(0) = m11
            sArray(1) = m22
            sArray(2) = m33
            sArray(3) = m12
            sArray(4) = m13
            sArray(5) = m23
            'calculate eigevector, convert to angles
            pArray = Char_cisla_prum_smer(sArray)
            gArray = Statistics(sArray)
            'Alfa = (pArray(0) + 180) Mod 360
            Alfa = pArray(0)
            Fi = 90 - pArray(1)
            'display angles on the form
            TextBox3.Text = Alfa
            TextBox6.Text = Fi
            GamaTxtBox.Text = Round(gArray(3), 3)
            KsTxtBox.Text = Round(gArray(4), 3)

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            'after the "plot axis as graphics" button is used
            'axis must be calculated first
            If TextBox4.Text = "" Then
                MsgBox("Calculate the axis first!")
                Exit Sub
            End If

            'read angles
            Dim Alfa, Fi As Double
            Alfa = TextBox4.Text
            Fi = TextBox5.Text

            'coint vectorcomponents
            Const PI = 3.14159265
            Dim osavektorX, osavektorY, osavektorZ As Double
            'MsgBox Alfa & Fi
            osavektorX = (Cos(Alfa * PI / 180)) * (Cos(Fi * PI / 180))
            osavektorY = (Sin(Alfa * PI / 180)) * (Cos(Fi * PI / 180))
            osavektorZ = Sin(Fi * PI / 180)
            'MsgBox osavektorX & osavektorY & osavektorZ

            'calculate xy coords in the diagram DF
            Dim Ro, yd, xd, x, y As Double
            Const r = 200
            Const xo = 200
            Const yo = 200

            Dim pPoint As IPoint
            pPoint = New Point
            Ro = r * Sqrt(1 - osavektorZ)
            yd = r * osavektorX / Sqrt(1 + osavektorZ)
            xd = r * osavektorY / Sqrt(1 + osavektorZ)
            x = xo + xd
            y = yo + yd
            'MsgBox "Ro je: " & Ro & "yd je: " & yd & "xd je: " & xd
            pPoint.X = x
            pPoint.Y = y
            'create graphic
            Dim pElement As IElement
            Dim pGraphicsContainer As IGraphicsContainer
            Dim pSymbol As IMarkerSymbol
            Dim pMElement As IMarkerElement
            pElement = New MarkerElement
            pElement.Geometry = pPoint
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            pGraphicsContainer = pMxDoc.FocusMap
            pGraphicsContainer.AddElement(pElement, 0)

            'symbol
            pSymbol = New SimpleMarkerSymbol

            Dim pCervena As IRgbColor
            pCervena = New RgbColor
            pCervena.Red = 169
            pCervena.Green = 0
            pCervena.Blue = 230
            pSymbol.Color = pCervena
            pSymbol.Size = 3
            pMElement = pElement
            pMElement.Symbol = pSymbol

            'refresh
            Dim pActiveView As IActiveView
            pActiveView = pMxDoc.FocusMap
            pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            '"create new shp and plot" button is used:
            'foldaxis must be calculated first
            If TextBox4.Text = "" Then
                MsgBox("Calculate the axis first!")
                Exit Sub
            End If
            'check inputs
            If ComboBox1.Text = "" Then
                MsgBox("The input layer must be chosen!")
                Exit Sub
            End If
            If ComboBox4.Text = "" Or ComboBox5.Text = "" Then
                MsgBox("Choose fields names of the input layer!")
                Exit Sub
            End If
            If TextBox7.Text = "" Or TextBox8.Text = "" Then
                MsgBox("Choose output folder and type a name for the output shapefile!")
                Exit Sub
            End If
            'read map
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap

            Dim pNewFLayer As IFeatureLayer
            Dim pNewFClass As IFeatureClass
            Dim pNewFeature As IFeature

            'toto řeší, že pokud během vynášení se změní aktivní DF, tak vstupní vrstvu, která je parametr pro tvorbu nového SHP (kopíruje jeho strukturu) nenajde
            'v aktivním DF a zahlásí Error. Proto, v následujícím cyklu ji hledá v aktivním DF, když ji najde i = 999, pokud i není 999 znamená to, že jsme v jiném DF
            'a přeruší se procedura.

            'check if the right DF is active
            Dim aLName As String
            Dim I, nl As Long
            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name
                If aLName = ComboBox1.Text Then
                    nl = 999
                    Exit For
                End If
            Next
            If Not nl = 999 Then
                MsgBox("The chosen input layer is in Data frame, which is currently NOT active. Chose layer from active Data Frame.")
                Exit Sub
            End If

            'pokud je zatržena možnost 2,vytvoří nový shp
            'zjistí spatial reference vstupní vrstvy, výstup bude mít stejný souř.systém
            'if "create new shp" option is selected,
            'read spatial ref of input layer
            Dim pSpatialReference As ISpatialReference
            Dim geoDataset As IGeoDataset
            geoDataset = pMap.Layer(I)
            pSpatialReference = geoDataset.SpatialReference

            'call Geoprocessing tool
            Dim pGP As IGeoProcessor
            pGP = New GeoProcessor
            pGP.OverwriteOutput = True

            'parameters array for "CreateFeatureClass" tool
            Dim pParameterArray As IVariantArray
            pParameterArray = New VarArray

            pParameterArray.Add(TextBox8.Text)
            pParameterArray.Add(TextBox7.Text)
            pParameterArray.Add("Point")
            pParameterArray.Add(ComboBox1.Text)
            'pokud nemá vstupní vrstva sdefinovaný souř. systém, nebude mít ani výsledná
            'If Not pSpatialReference.FactoryCode = 0 Then
            'pGP.SetEnvironmentValue("outputCoordinateSystem", pSpatialReference.FactoryCode)
            'End If
            'CS is set, pevně bude nastaven Křovák pro nový shp, aby výsledek spadal do kruhu diagramu
            '7.12.2014 ruším souř.systém, výsledek nemá ss. Křovak byl matoucí.
            'pGP.SetEnvironmentValue("outputCoordinateSystem", 102067)

            'create new shp
            Dim pResult As IGeoProcessorResult
            pResult = pGP.Execute("CreateFeatureClass_management", pParameterArray, Nothing)

            'read new layer, pNewFeature
            pNewFLayer = pMap.Layer(0)
            pNewFClass = pNewFLayer.FeatureClass

            'table, fields
            Dim pTable As ITable
            pTable = pNewFClass

            Dim pFields As IFields
            pFields = pTable.Fields

            Dim smers, sklons, smerl, sklonl, styp, ltyp As String
            smers = ComboBox4.Text
            sklons = ComboBox5.Text
            smerl = ComboBox7.Text
            sklonl = ComboBox8.Text
            styp = ComboBox3.Text
            ltyp = ComboBox6.Text

            Dim ityp, iAz, iSk As Integer
            ityp = 0
            iAz = 0
            iSk = 0

            'check input fields names
            If Not smers = "" Then
                iAz = pFields.FindField(smers)
                iSk = pFields.FindField(sklons)
                ityp = pFields.FindField(styp)
                If iAz = -1 Or iSk = -1 Or ityp = -1 Then
                    MsgBox("Some of the fields weren´t found in the input table.")
                    Exit Sub
                End If
            End If

            If Not smerl = "" Then
                iAz = pFields.FindField(smerl)
                iSk = pFields.FindField(sklonl)
                ityp = pFields.FindField(ltyp)
                If iAz = -1 Or iSk = -1 Or ityp = -1 Then
                    MsgBox("Some of the fields weren´t found in the input table.")
                    Exit Sub
                End If
            End If

            'calculate coordinates in the diagram DF
            Const PI = 3.14159265
            Dim Alfa, Fi As Double
            Alfa = TextBox4.Text
            Fi = TextBox5.Text
            Dim osavektorX, osavektorY, osavektorZ As Double
            'MsgBox Alfa & Fi

            osavektorX = (Cos(Alfa * PI / 180)) * (Cos(Fi * PI / 180))
            osavektorY = (Sin(Alfa * PI / 180)) * (Cos(Fi * PI / 180))
            osavektorZ = Sin(Fi * PI / 180)
            'MsgBox osavektorX & osavektorY & osavektorZ

            Dim Ro, yd, xd, x, y As Double
            Const r = 200
            Const xo = 200
            Const yo = 200

            Dim pPoint As IPoint
            pPoint = New Point

            Ro = r * Sqrt(1 - osavektorZ)

            yd = r * osavektorX / Sqrt(1 + osavektorZ)
            xd = r * osavektorY / Sqrt(1 + osavektorZ)
            x = xo + xd
            y = yo + yd
            'MsgBox "Ro je: " & Ro & "yd je: " & yd & "xd je: " & xd
            pPoint.X = x
            pPoint.Y = y

            'create new feature and fill attributes
            pNewFeature = pNewFClass.CreateFeature
            pNewFeature.Shape = pPoint
            pNewFeature.Value(iAz) = TextBox4.Text
            pNewFeature.Value(iSk) = TextBox5.Text
            pNewFeature.Value(ityp) = "axis"
            pNewFeature.Store()
            'refresh
            Dim pActiveView As IActiveView
            pActiveView = pMxDoc.FocusMap
            pMxDoc.ActiveView.Refresh()
        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try
            'after "add axis to existing shp: option is selected
            'axis must be calculated first
            If TextBox4.Text = "" Then
                MsgBox("Calculate the axis first!")
                Exit Sub
            End If
            'open form
            Dim frm As New OsaPlotForm
            frm.Show()
            frm.TextBox1.Text = TextBox4.Text
            frm.TextBox2.Text = TextBox5.Text

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub Button11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        'Plot as arcs
        Try
            'check inputs
            If RadioButton3.Checked = True Then
                If ComboBox9.Text = "" Then
                    MsgBox("Choose an existing SHP where the plotted points will be added!")
                    Exit Sub
                End If
            End If
            If RadioButton2.Checked = True And (TextBox1.Text = "" Or TextBox2.Text = "") Then
                MsgBox("Select the output workspace (folder) and type a name for the output shapefile!")
                Exit Sub
            End If
            'ComboBox1 je pro vstupní vrstvu, musí být vybrána. Pokud není, zpráva a konec.
            If ComboBox1.Text = "" Then
                MsgBox("Choose an input layer!")
                Exit Sub
            End If
            'ComboBox2 - field for control point ID
            If ComboBox2.Text = "" Or ComboBox3.Text = "" Or ComboBox4.Text = "" Or ComboBox5.Text = "" Then
                MsgBox("Some of the fields for S structure are not filled!")
                Exit Sub
            End If

            'read map, layer
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap
            Dim aLName As String
            Dim I As Long
            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name
                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            'input layer must not be raster layer
            If TypeOf pMap.Layer(I) Is IRasterLayer Then
                MsgBox("You chose a raster layer.. but we need vector point feature class or shapefile!")
                Exit Sub
            End If

            Dim pFLayer As IFeatureLayer
            pFLayer = pMap.Layer(I)

            'input layer must have data source
            If pFLayer.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If
            'input layer must not be of a sde data source
            If pFLayer.DataSourceType = "SDE Feature Class" Then
                MsgBox("Input layer is a SDE Feature Class. Export it to shapefile or database Feature Class first")
                Exit Sub
            End If

            'read table, fields
            Dim pFClass As IFeatureClass = Nothing
            pFClass = pFLayer.FeatureClass

            Dim pTable As ITable = Nothing
            pTable = pFClass

            Dim pFields As IFields = Nothing
            pFields = pTable.Fields

            'create "N" field, hat will be used for link

            'Dim islN As Integer = Nothing
            'islN = pFields.FindField("N")

            'If islN = -1 Then
            ''tento sloupec je pro identifikaci každého měřeného prvku, abyse potom mohl ztotožnit s vyneseným prvkem v projekci a zpátky.
            'Dim pFieldEditN As IFieldEdit
            'pFieldEditN = New Field
            'pFieldEditN.Name_2 = "N"
            'pFieldEditN.Type_2 = esriFieldType.esriFieldTypeInteger
            'pFClass.AddField(pFieldEditN)
            'End If

            Dim SelYes As Integer = Nothing
            'raed options in the form
            If CheckBox1.Checked = True Then
                SelYes = 1
            Else
                SelYes = 0
            End If

            '"create new shp" option
            Dim shpYes As Integer = Nothing
            If RadioButton2.Checked = True Then
                shpYes = 1
            End If
            If RadioButton1.Checked = True Then
                MsgBox("Arc cannot be plotted as a graphics, create new shapefile!")
                Exit Sub
            End If

            '"existing shp" option
            If RadioButton3.Checked = True Then
                shpYes = 2
            End If

            'read selection
            Dim pFSelection As IFeatureSelection = Nothing
            pFSelection = pFLayer
            Dim pInSelSet As ISelectionSet = Nothing
            pInSelSet = pFSelection.SelectionSet

            Dim cbod, smers, sklons, styp, smerl, sklonl, ltyp, filePath, name, Par4 As String
            cbod = ComboBox2.Text
            smers = ComboBox4.Text
            sklons = ComboBox5.Text
            smerl = ComboBox7.Text
            sklonl = ComboBox8.Text
            styp = ComboBox3.Text
            ltyp = ComboBox6.Text
            filePath = TextBox1.Text
            name = TextBox2.Text
            Par4 = ComboBox1.Text
            Dim VeX_s, VeY_s, VeZ_s As Double
            'Dim VeX_l, VeY_l, VeZ_l As Double

            Const r = 200
            Const xo = 200
            Const yo = 200

            Dim Ro, xd, yd, x, y As Double
            Dim ieN, iAZ_S, iSk_S, iAZ_L, iSk_L, iStyp, iLtyp, iDB As Integer
            ieN = 0

            iAZ_S = 0
            iSk_S = 0
            iAZ_L = 0
            iSk_L = 0
            iStyp = 0
            iLtyp = 0
            iDB = 0

            Dim pFeature As IFeature = Nothing
            iDB = pFields.FindField(cbod)
            'ieN = pFields.FindField("N")
            If Not iDB > 0 Then
                MsgBox("Field " & cbod & " wasn´t found int the input table")
                Exit Sub
            End If

            'indexes of the fields
            If Not smers = "" Then
                iAZ_S = pFields.FindField(smers)
                iSk_S = pFields.FindField(sklons)
                iStyp = pFields.FindField(styp)
                If iAZ_S = -1 Or iSk_S = -1 Or iStyp = -1 Then
                    MsgBox("Some of the fields for S structure weren´t found")
                    Exit Sub
                End If
            End If

            If Not smerl = "" Then
                iAZ_L = pFields.FindField(smerl)
                iSk_L = pFields.FindField(sklonl)
                iLtyp = pFields.FindField(ltyp)
                If iAZ_L = -1 Or iSk_L = -1 Or iLtyp = -1 Then
                    MsgBox("Some of the fields for L structure weren´t found")
                    Exit Sub
                End If
            End If

            'Pokud nejsou zadané sloupce pro s prvky nebo l prvky, je iStyp.... 0.
            'V tomto ppřípadě vůbec nezačne cyklus počítáni vektorů
            'Pokud je něco zadané a on to pak nenajde je iStyp.... -1.
            'V tomto případě následuje konec procedury:

            If iAZ_S = -1 Then
                Exit Sub
            End If
            If iAZ_L = -1 Then
                Exit Sub
            End If

            'cursor, read the first feature in the table
            Dim pFCursor As IFeatureCursor = Nothing
            'Set pFCursor = pFClass.Update(Nothing, True)
            'předchozí řádek plus UpdateFeature místo pFeature.Store použiju, pokud chci vynášet všechny měření, ne jen selection

            If SelYes = 1 Then
                pInSelSet.Search(Nothing, True, pFCursor)
            Else
                pFCursor = pFClass.Update(Nothing, True)
            End If

            'Dim pElement As IElement = Nothing
            'Dim pGraphicsContainer As IGraphicsContainer = Nothing
            'Dim pSymbol As IMarkerSymbol = Nothing
            'Dim pMElement As IMarkerElement = Nothing
            Dim pNewFLayer As IFeatureLayer = Nothing
            Dim pNewFClass As IFeatureClass = Nothing
            Dim pNewFeature As IFeature = Nothing
            Dim pFieldEditN As IFieldEdit = Nothing
            Dim sArray(1) As Integer
            Dim lArray(1) As Integer
            Dim kArray(2) As Double
            Dim mArray(2) As Double

            Dim pPointA As IPoint = Nothing
            Dim pPointB As IPoint = Nothing
            Dim pPointC As IPoint = Nothing

            If shpYes = 1 Then
                'Dim pSpatialReference As ISpatialReference
                'Dim geoDataset As IGeoDataset
                'geoDataset = pFLayer
                'pSpatialReference = geoDataset.SpatialReference

                Try
                    'pokud je zatržena možnost vytvořit nový shp /shpYes = 1
                    'zavolá Geoprocessing tool
                    Dim pGP As IGeoProcessor
                    pGP = New GeoProcessor
                    pGP.OverwriteOutput = True

                    'vytvoří matici parametrů pro nástroj "CreateFeatureClass"
                    Dim pParameterArray As IVariantArray
                    pParameterArray = New VarArray

                    pParameterArray.Add(filePath)
                    pParameterArray.Add(name)
                    pParameterArray.Add("Polyline")
                    pParameterArray.Add(Par4)
                    'pokud nemá vstupní vrstva sdefinovaný souř. systém, nebude mít ani výsledná
                    'If Not pSpatialReference.FactoryCode = 0 Then
                    'pGP.SetEnvironmentValue("outputCoordinateSystem", pSpatialReference.FactoryCode)
                    'End If
                    'pevně nastavuju Křováka pro nový shp, aby výsledek spadal do kruhu diagramu.
                    '7.12.2014 ruším souř.systém, výsledek nemá ss. Křovak byl matoucí.
                    'pGP.SetEnvironmentValue("outputCoordinateSystem", 102067)

                    'vytvoř novou třídu prvků
                    Dim pResult As IGeoProcessorResult
                    pResult = pGP.Execute("CreateFeatureClass_management", pParameterArray, Nothing)

                Catch ex As Exception
                    MsgBox("error creating new shapefile. " & Err.Number & ": " & Err.Description)
                    Exit Sub
                End Try

                'načte novou vrstvu a nový pNewFeature
                Dim aLName3 As String
                Dim zz As Long

                For zz = 0 To pMap.LayerCount - 1
                    aLName3 = pMap.Layer(zz).Name

                    If aLName3 = TextBox2.Text Then
                        Exit For
                    End If
                Next

                pNewFLayer = pMap.Layer(zz)
                pNewFClass = pNewFLayer.FeatureClass

                'add "N" field
                Dim islN As Integer = Nothing
                islN = pFields.FindField("N")

                If islN = -1 Then
                    pFieldEditN = New Field
                    pFieldEditN.Name_2 = "N"
                    pFieldEditN.Type_2 = esriFieldType.esriFieldTypeInteger
                    pNewFClass.AddField(pFieldEditN)

                End If

            ElseIf shpYes = 2 Then
                'do pNewFLayer se načte vybraná vrstva, do které se výsledek má přidat
                'Dim pNewFLayer As IFeatureLayer
                'Dim aLName As String
                Dim l As Long

                For l = 0 To pMap.LayerCount - 1
                    aLName = pMap.Layer(l).Name

                    If aLName = ComboBox9.Text Then
                        Exit For
                    End If
                Next

                pNewFLayer = pMap.Layer(l)
                pNewFClass = pNewFLayer.FeatureClass
                If Not pNewFLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
                    MsgBox("You haven´t chosen polyline layer to add arcs. Arcs are polyline geometry, choose polyline shapefile!")
                    Exit Sub
                End If
            End If
            pFeature = pFCursor.NextFeature
            Do Until pFeature Is Nothing
                If Not iStyp = 0 Then
                    'PointA_________________________¨spádnicový bod
                    sArray(0) = pFeature.Value(iAZ_S)
                    sArray(1) = pFeature.Value(iSk_S)

                    kArray = Vektory_lineace(sArray)
                    VeX_s = kArray(0)
                    VeY_s = kArray(1)
                    VeZ_s = kArray(2)

                    'další výpočty na každém měření vypočte ro, xd a yd, přičte k xo a yo
                    ' vytvořím point geometry a vynesem
                    Ro = r * (1 - VeZ_s) ^ (1 / 2)
                    yd = r * VeX_s / (1 + VeZ_s) ^ (1 / 2)
                    xd = r * VeY_s / (1 + VeZ_s) ^ (1 / 2)
                    x = xo + xd
                    y = yo + yd
                    pPointA = New Point
                    pPointA.X = x
                    pPointA.Y = y

                    'PointB_____________velká kružnice bude procházet skrz pointB a pointC
                    Dim alfaB As Double
                    alfaB = (pFeature.Value(iAZ_S) + 270) Mod 360

                    sArray(0) = alfaB
                    sArray(1) = 0
                    kArray = Vektory_lineace(sArray)
                    VeX_s = kArray(0)
                    VeY_s = kArray(1)
                    VeZ_s = kArray(2)
                    Ro = r * (1 - VeZ_s) ^ (1 / 2)
                    yd = r * VeX_s / (1 + VeZ_s) ^ (1 / 2)
                    xd = r * VeY_s / (1 + VeZ_s) ^ (1 / 2)
                    x = 0
                    y = 0
                    x = xo + xd
                    y = yo + yd

                    pPointB = New Point
                    pPointB.X = x
                    pPointB.Y = y

                    'PointC_________________________________________
                    Dim alfaC As Double
                    alfaC = (pFeature.Value(iAZ_S) + 90) Mod 360

                    sArray(0) = alfaC
                    sArray(1) = 0
                    kArray = Vektory_lineace(sArray)
                    VeX_s = kArray(0)
                    VeY_s = kArray(1)
                    VeZ_s = kArray(2)

                    Ro = r * (1 - VeZ_s) ^ (1 / 2)
                    yd = r * VeX_s / (1 + VeZ_s) ^ (1 / 2)
                    xd = r * VeY_s / (1 + VeZ_s) ^ (1 / 2)
                    x = 0
                    y = 0
                    x = xo + xd
                    y = yo + yd

                    pPointC = New Point
                    pPointC.X = x
                    pPointC.Y = y

                    If Not pFeature.Value(iSk_S) = 90 Then
                        'center point________________________
                        'nejprv vypočítám poloměr velké kružnice rr,pak d což je vzdálenost spádnicového bodu oblouku
                        'od středu diagramu, xx je rozdíl - vzdálenost mezi středem velkékružnice a středem diagramu
                        'pak můžu vypočítat souřadnice středu velkého diagramu, resp. vydál od počátku souř.s. 0,0.
                        'které přičtu ke tředu diagramu (200, 200).
                        'azimut sklonu asklonpřevedu ze stupní na radiány, F je součást vzorců, tak pro přehlednost vypoč. zvlášť.
                        'výpočet rr a d je pro jednotkovou kružnici, takže násobím 200(poloměr diagramu)

                        Dim Alfa, Fi, F As Double
                        Const PI = 3.1416
                        Alfa = pFeature.Value(iAZ_S) * PI / 180
                        Fi = pFeature.Value(iSk_S) * PI / 180
                        F = (PI / 4) - (Fi / 2)
                        Dim centerPoint As IPoint = New Point
                        Dim rr, d, xx As Double
                        rr = (2 ^ (1 / 2) * Sin(F) + Sin(Fi) / (2 * (2 ^ (1 / 2)) * Sin(F))) * 200
                        d = (2 * Sin(F) / (2 ^ (1 / 2))) * 200
                        xx = rr - d

                        x = 0
                        y = 0
                        x = xx * Cos(Alfa + PI / 2)
                        y = -xx * Sin(Alfa + PI / 2)

                        x = (xo + x)
                        y = (yo + y)

                        centerPoint.X = x
                        centerPoint.Y = y

                        'If shpYes = 1 Or shpYes = 2 Then
                        'create arc from:  centerPoint, fromPoint and ToPoint
                        Dim circularArc As ICircularArc = New CircularArc
                        circularArc.PutCoords(centerPoint, pPointB, pPointC, ESRI.ArcGIS.Geometry.esriArcOrientation.esriArcClockwise)
                        Dim pSegCol As ISegmentCollection
                        pSegCol = New Polyline
                        pSegCol.AddSegment(circularArc)

                        'create polyline feature
                        pNewFeature = pNewFClass.CreateFeature
                        pNewFeature.Shape = pSegCol
                        pNewFeature.Value(pNewFeature.Fields.FindField(ComboBox3.Text)) = pFeature.Value(iStyp)
                        pNewFeature.Value(pNewFeature.Fields.FindField(ComboBox4.Text)) = pFeature.Value(iAZ_S)
                        pNewFeature.Value(pNewFeature.Fields.FindField(ComboBox5.Text)) = pFeature.Value(iSk_S)
                        pNewFeature.Value(pNewFeature.Fields.FindField(ComboBox2.Text)) = pFeature.Value(iDB)
                        pNewFeature.Value(pNewFeature.Fields.FindField("N")) = pFeature.Value(0)
                        pNewFeature.Store()
                    Else
                        'if dip is 90deg, instead of an arc, straight line is drawn, perpendiclar to dipdir
                        'line from pointB to pointC
                        'point collection
                        Dim pPointCollection As IPointCollection
                        pPointCollection = New Polyline
                        pPointCollection.AddPoint(pPointB)
                        pPointCollection.AddPoint(pPointC)
                        'create line
                        pNewFeature = pNewFLayer.FeatureClass.CreateFeature
                        pNewFeature.Shape = pPointCollection
                        pNewFeature.Value(pNewFeature.Fields.FindField(ComboBox3.Text)) = pFeature.Value(iStyp)
                        pNewFeature.Value(pNewFeature.Fields.FindField(ComboBox4.Text)) = pFeature.Value(iAZ_S)
                        pNewFeature.Value(pNewFeature.Fields.FindField(ComboBox5.Text)) = pFeature.Value(iSk_S)
                        pNewFeature.Value(pNewFeature.Fields.FindField(ComboBox2.Text)) = pFeature.Value(iDB)
                        pNewFeature.Value(pNewFeature.Fields.FindField("N")) = pFeature.Value(0)
                        pNewFeature.Store()
                    End If
                End If
                pFeature = pFCursor.NextFeature
            Loop
            Dim pActiveView As IActiveView
            pActiveView = pMxDoc.FocusMap
            pActiveView.Refresh()
            pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    'Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    'Dim pMxDoc As IMxDocument
    'pMxDoc = My.ArcMap.Document
    'Dim pMap As IMap
    'pMap = pMxDoc.FocusMap

    'Dim pNewFLayer As IFeatureLayer
    'Dim pNewFClass As IFeatureClass
    'Dim pNewFeature As IFeature

    'pNewFLayer = pMap.Layer(0)
    'pNewFClass = pNewFLayer.FeatureClass
    ''create arc from:  centerPoint, fromPoint and ToPoint

    ' Dim centerPoint As IPoint = New Point
    'Dim fromPoint As IPoint = New Point
    'Dim toPoint As IPoint = New Point
    'centerPoint.X = 0
    'centerPoint.Y = 0
    'fromPoint.X = 100
    'fromPoint.Y = 100
    'toPoint.X = -100
    'toPoint.Y = -100
    'MsgBox("tu")
    'Dim circularArc As ICircularArc = New CircularArc
    'circularArc.PutCoords(centerPoint, fromPoint, toPoint, ESRI.ArcGIS.Geometry.esriArcOrientation.esriArcClockwise)
    'MsgBox("tu")
    'Dim pSegCol As ISegmentCollection
    'pSegCol = New Polyline
    'pSegCol.AddSegment(circularArc)
    'MsgBox("tu")
    ''create polyline feature
    'pNewFeature = pNewFClass.CreateFeature
    'MsgBox("tu")
    'pNewFeature.Shape = pSegCol
    'MsgBox("tu")
    'pNewFeature.Store()
    'End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        'open help form
        Dim frm As New HelpForm
        frm.Show()
    End Sub

    Private Sub Button12_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Try
            'folder browser
            Dim FolderBrowserDialog1 As New FolderBrowserDialog
            With FolderBrowserDialog1

                '.RootFolder = Environment.SpecialFolder.MyDocuments
                .SelectedPath = "c:"
                .Description = "Select folder"

                If .ShowDialog = DialogResult.OK Then
                    TextBox8.Text = (.SelectedPath)
                End If
            End With

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try
    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox3.Enter

    End Sub
End Class

