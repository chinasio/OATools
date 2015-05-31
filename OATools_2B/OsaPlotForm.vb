Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports System.Math
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry


Public Class OsaPlotForm

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        'after the input layer is chosen, read the map and the input layer
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

            pLDB = pMap.Layer(I)
            'input layer must not be a raster layer
            If TypeOf pLDB Is IRasterLayer Then
                MsgBox("You chose a raster layer.. but we need vector point feature class or shapefile!")
                Exit Sub
            End If
            'input layer must have data source
            If pLDB.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            pFLDB = pMap.Layer(I)
            'input layer must not be of sde data source
            If pFLDB.DataSourceType = "SDE Feature Class" Then
                MsgBox("Input layer is a SDE Feature Class, which is not editable. Export it to shapefile first.")
                Exit Sub
            End If

            'table, fields
            Dim pFClassDB As IFeatureClass
            pFClassDB = pFLDB.FeatureClass
            Dim pTableDB As ITable
            pTableDB = pFClassDB
            Dim pFieldsDB As IFields
            pFieldsDB = pTableDB.Fields
            
            ComboBox3.Items.Add("")
            ComboBox4.Items.Add("")
            ComboBox5.Items.Add("")

            'add fields name to comboboxes
            Dim j As Long
            For j = 0 To pFieldsDB.FieldCount - 1
                If Not pFieldsDB.Field(j).Type = esriFieldType.esriFieldTypeOID And Not pFieldsDB.Field(j).Type = esriFieldType.esriFieldTypeGeometry Then
                    ComboBox3.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox4.Items.Add(pFieldsDB.Field(j).Name)
                    ComboBox5.Items.Add(pFieldsDB.Field(j).Name)
                End If
            Next

            'automatically fill the field names if recommended structure
            Dim L1, L2, L3 As Long
            L1 = pTableDB.FindField("L_TYPE")
            If L1 > (-1) Then
                ComboBox3.Text = "L_TYPE"
            End If

            L2 = pTableDB.FindField("TREND")
            If L2 > (-1) Then
                ComboBox4.Text = "TREND"
            End If

            L3 = pTableDB.FindField("PLUNGE")
            If L3 > (-1) Then
                ComboBox5.Text = "PLUNGE"
            End If
            Exit Sub

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            'load field names if already saved
            'read from text file
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
                'fill comboboxes
                While Not MyReader.EndOfData

                    CurrentRow = MyReader.ReadFields()

                    ComboBox3.Text = CurrentRow(2)
                    ComboBox4.Text = CurrentRow(0)
                    ComboBox5.Text = CurrentRow(1)
                End While
            End Using

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            'after click on "add axis to shp and plot" button
            'input layer must be chosen
            If ComboBox1.Text = "" Then
                MsgBox("Choose an input layer!")
                Exit Sub
            End If

            'read the map
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document
            Dim pMap As IMap
            pMap = pMxDoc.FocusMap

            Dim pNewFLayer As IFeatureLayer
            Dim pNewFClass As IFeatureClass
            Dim pNewFeature As IFeature

            'read the input layer
            Dim aLName As String
            Dim l As Long
            For l = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(l).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            pNewFLayer = pMap.Layer(l)
            pNewFClass = pNewFLayer.FeatureClass

            'input layer must have data source
            If pNewFLayer.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If
            'input layer must be of point geometry
            If Not pNewFLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                MsgBox("You haven´t chosen point layer to add axis. Axis is point geometry, choose point shapefile!")
                Exit Sub
            End If

            'table, fields
            Dim pTable As ITable
            pTable = pNewFClass

            Dim pFields As IFields
            pFields = pTable.Fields

            Dim iAz, iSk, iStyp As Integer

            'check fields
            iAz = pFields.FindField(ComboBox4.Text)
            iSk = pFields.FindField(ComboBox5.Text)
            iStyp = pFields.FindField(ComboBox3.Text)
            If iAz < 1 Or iSk < 1 Or iStyp < 1 Then
                MsgBox("Some of the fields weren´t found in the input table!")
                Exit Sub
            End If

            'plot orientation of fold axis
            Const PI = 3.14159265
            Dim Alfa, Fi As Double
            Alfa = TextBox1.Text
            Fi = TextBox2.Text
            'vector components
            Dim osavektorX, osavektorY, osavektorZ As Double
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
            'xy coordinates in the diagram DF
            yd = r * osavektorX / Sqrt(1 + osavektorZ)
            xd = r * osavektorY / Sqrt(1 + osavektorZ)
            x = xo + xd
            y = yo + yd
            'MsgBox "Ro je: " & Ro & "yd je: " & yd & "xd je: " & xd
            pPoint.X = x
            pPoint.Y = y

            'create new feature in the selected layer, fill attributes
            pNewFeature = pNewFClass.CreateFeature
            pNewFeature.Shape = pPoint
            pNewFeature.Value(iAz) = TextBox1.Text
            pNewFeature.Value(iSk) = TextBox2.Text
            pNewFeature.Value(iStyp) = "axis"
            pNewFeature.Store()

            'refresh
            Dim pActiveView As IActiveView
            pActiveView = pMxDoc.FocusMap
            pMxDoc.ActiveView.Refresh()

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub OsaPlotForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'after the form is loaded, read the map
            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            Dim aLName As String
            Dim I As Long
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap

            'read all layers in TOC to combobox
            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name
                ComboBox1.Items.Add(aLName)
            Next

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
            Exit Sub
        End Try
    End Sub
End Class