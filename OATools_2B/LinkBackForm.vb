Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

Public Class LinkBackForm

    Private Sub LinkBackForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'after the form loads, read the map and layers into combobox
            Dim pMxDoc As IMxDocument
            Dim pMap As IMap
            Dim aLName As String
            Dim I As Long
            pMxDoc = My.ArcMap.Document
            pMap = pMxDoc.FocusMap
            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name
                ComboBox1.Items.Add(aLName)
                ComboBox2.Items.Add(aLName)
            Next

            'ComboBox1.Text = pMap.Layer(0).name
            'ComboBox2.Text = pMap.Layer(1).name

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            'after the input layer is selected, read the map and the input layer
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

            'input layer mustn´t be a raster layer
            If TypeOf pMap.Layer(I) Is IRasterLayer Then
                MsgBox("You chose a raster layer.. but we need vector point feature class or shapefile!")
                Exit Sub
            End If

            Dim pFLDB As IFeatureLayer
            pFLDB = pMap.Layer(I)

            'input layer must have data source
            If pFLDB.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            'table and fields
            Dim pFClassDB As IFeatureClass
            pFClassDB = pFLDB.FeatureClass

            Dim pTableDB As ITable
            pTableDB = pFClassDB
            Dim pFieldsDB As IFields
            pFieldsDB = pTableDB.Fields
            'Dim pFieldDB As IField
            'read fields and display incombobox
            'Dim j As Long
            'For j = 0 To pFieldsDB.FieldCount - 1
            'ComboBox3.Items.Add(pFieldsDB.Field(j).Name)
            'Next
        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            'click on "Link diagram to map" button

            'From and To layers must be set
            If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
                MsgBox("Choose FROM and TO layers!")
                Exit Sub
            End If
           
            'read the map
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document

            Dim pMap As IMap
            pMap = pMxDoc.FocusMap

            'read layer(i), it is what we want to link with something, (the layer with points in the diagram)
            Dim aLName As String
            Dim I As Long
            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            'pFLayer0 is the layer with points plotted in diagram
            Dim pFLayer0 As IFeatureLayer
            pFLayer0 = pMap.Layer(I)
            'input layer must have data source
            If pFLayer0.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If
            Dim pFClass0 As IFeatureClass
            pFClass0 = pFLayer0.FeatureClass

            'read layer(j), the layer we want to link with (original tectonic layer in the map)
            Dim j As Long
            For j = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(j).Name
                If aLName = ComboBox2.Text Then
                    Exit For
                End If
            Next

            'pFLayer1 is the layer with original data
            Dim pFLayer1 As IFeatureLayer
            pFLayer1 = pMap.Layer(j)
            'input layer must have data source
            If pFLayer1.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            Dim pFClass1 As IFeatureClass
            pFClass1 = pFLayer1.FeatureClass

            'read selection
            Dim pFSelection0 As IFeatureSelection
            pFSelection0 = pFLayer0
            Dim pInSelSet0 As ISelectionSet
            pInSelSet0 = pFSelection0.SelectionSet

            Dim pFSelection1 As IFeatureSelection
            pFSelection1 = pFLayer1
            'Dim pInSelSet1 As ISelectionSet
            'Set pInSelSet1 = pFSelection1.SelectionSet

            Dim pTable As ITable
            pTable = pFClass0

            Dim pFields As IFields
            pFields = pTable.Fields


            'plot tool can plot in a new shp, that has the same table structure as the original table
            'plot tool adds an "N" field to the new table.
            '"N" field is populated by FID´s from the original table
            'Link tool use "N" field to link between original and the new table

            Dim ieN, ieF As Integer
            Dim pFeature0 As IFeature
            Dim pFeature1 As IFeature
            'N field is located
            ieN = pFields.FindField("N")
            If ieN = -1 Then
                MsgBox("N fields wasn't found in the FROM layer.")
                Exit Sub
            End If
            ieF = pFields.FindField("FID")

            Dim pFCursor0 As IFeatureCursor = Nothing

            'create cursor and read the first feature
            Dim pFCursor1 As IFeatureCursor
            pFCursor1 = pFClass1.Update(Nothing, True)
            pFeature1 = pFCursor1.NextFeature

            'select corresponding features 
            Do Until pFeature1 Is Nothing

                pInSelSet0.Search(Nothing, False, pFCursor0)
                pFeature0 = pFCursor0.NextFeature
                Do Until pFeature0 Is Nothing

                    If pFeature1.Value(ieF) = pFeature0.Value(ieN) Then
                        pFSelection1.Add(pFeature1)
                    End If
                    pFeature0 = pFCursor0.NextFeature
                Loop

                pFeature1 = pFCursor1.NextFeature
            Loop
            With pMxDoc
                .ActiveView.Refresh()
                .UpdateContents()
            End With

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'click on "Link map to diagram" button

        Try
            
            'read the map
            Dim pMxDoc As IMxDocument
            pMxDoc = My.ArcMap.Document

            Dim pMap As IMap
            pMap = pMxDoc.FocusMap

            'read layer(i),it is what we want to link to something (points with tectonic orientations in the map)
            Dim aLName As String
            Dim I As Long

            For I = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(I).Name

                If aLName = ComboBox1.Text Then
                    Exit For
                End If
            Next

            'pFLayer0 is layer with original data
            Dim pFLayer0 As IFeatureLayer
            pFLayer0 = pMap.Layer(I)
            'input layer must have data source
            If pFLayer0.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            Dim pFClass0 As IFeatureClass
            pFClass0 = pFLayer0.FeatureClass

            'read layer(j), it is what we link with (points in diagram)
            Dim j As Long
            For j = 0 To pMap.LayerCount - 1
                aLName = pMap.Layer(j).Name

                If aLName = ComboBox2.Text Then
                    Exit For
                End If
            Next

            ' pFLayer1 is the layer with points plotted in diagram
            Dim pFLayer1 As IFeatureLayer
            pFLayer1 = pMap.Layer(j)
            'input layer must have data source
            If pFLayer1.Valid = False Then
                MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
                Exit Sub
            End If

            Dim pFClass1 As IFeatureClass
            pFClass1 = pFLayer1.FeatureClass

            'read selection 
            Dim pFSelection0 As IFeatureSelection
            pFSelection0 = pFLayer0
            Dim pInSelSet0 As ISelectionSet
            pInSelSet0 = pFSelection0.SelectionSet
            Dim pFSelection1 As IFeatureSelection
            pFSelection1 = pFLayer1
            'Dim pInSelSet1 As ISelectionSet
            'Set pInSelSet1 = pFSelection1.SelectionSet

            'table, fields
            Dim pTable As ITable
            pTable = pFClass1
            Dim pFields As IFields
            pFields = pTable.Fields

            'plot tool can plot in a new shp, that has the same table structure as the original table
            'plot tool adds an "N" field to the new table.
            '"N" field is populated by FID´s from the original table
            'Link tool use "N" field to link between original and the new table


            Dim ieN, ieF As Integer
            'Dim iAz As Integer
            Dim pFeature0 As IFeature
            Dim pFeature1 As IFeature

            'N field is located
            ieN = pFields.FindField("N")
            If ieN = -1 Then
                MsgBox("N field wasn´t found in the TO layer.")
                Exit Sub
            End If
            ieF = pFields.FindField("FID")
           
            'create cursor and read the first feature
            Dim pFCursor0 As IFeatureCursor = Nothing
            Dim pFCursor1 As IFeatureCursor
            pFCursor1 = pFClass1.Update(Nothing, True)
            pFeature1 = pFCursor1.NextFeature

            'select corresponding features 
            Do Until pFeature1 Is Nothing
                pInSelSet0.Search(Nothing, False, pFCursor0)
                pFeature0 = pFCursor0.NextFeature
                Do Until pFeature0 Is Nothing
                    'MsgBox pFeature0.Value(ieN)

                    If pFeature1.Value(ieN) = pFeature0.Value(ieF) Then

                        pFSelection1.Add(pFeature1)
                    End If
                    pFeature0 = pFCursor0.NextFeature
                Loop
                pFeature1 = pFCursor1.NextFeature
            Loop
            With pMxDoc

                .ActiveView.Refresh()
                .UpdateContents()
            End With

        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description)
        End Try

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

        'after change in combobox for layer to be linked with, read the map
        Dim pMxDoc As IMxDocument
        pMxDoc = My.ArcMap.Document
        Dim pMap As IMap
        pMap = pMxDoc.FocusMap
        'read the input layer
        Dim aLName As String
        Dim I As Long
        For I = 0 To pMap.LayerCount - 1
            aLName = pMap.Layer(I).Name
            If aLName = ComboBox1.Text Then
                Exit For
            End If
        Next

        'input layer must´t be a raster layer
        If TypeOf pMap.Layer(I) Is IRasterLayer Then
            MsgBox("You chose a raster layer.. but we need vector point feature class or shapefile!")
            Exit Sub
        End If

        Dim pFLDB As IFeatureLayer
        pFLDB = pMap.Layer(I)

        'input layer must have data source
        If pFLDB.Valid = False Then
            MsgBox("Input layer is not valid! It has probably no data source. Choose another one!")
            Exit Sub
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'open help form
        Dim frm As New HelpForm
        frm.Show()
    End Sub
End Class