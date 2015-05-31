Public Class HelpForm
    Private Sub LinkLabel8_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        'New Projection help
        TextBox1.Text = "This button should be used first before the orientations of geological structures are plotted. " & _
        vbNewLine & "It is an azimuthal lower hemisphere equal-area projection. This projection is added to the map document as a new Data Frame." & _
        vbNewLine & "It is switched simultaneously to Layout View in order to see both the projection and the map." & _
        vbNewLine & "First time you use this button you are prompted to select folder, where the cross.shp is saved. It is in the downloaded unzipped file (OATools_2B_for_ArcGIS_10.2)." & _
        vbNewLine & "The cross.shp is the shape of the diagram and by adding it, the reference scale and the appropriate data frame size in map units (400 x 400) are set." & _
        vbNewLine & "The directory to this shapefile is saved for the further use (on Desktop, stereonet.txt)." & _
        vbNewLine & "Drag the layer with structural data to the Projection Data Frame and make sure it is active before you use the Plot Tool, Density Distributin, or Rose Histogram tools."
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        TextBox1.Text = "This tool plots selected features from the input layer to azimuthal equal-area projection (lower hemisphere)." & _
vbNewLine & "Create New Projection by the 'New Projection' button first. Make it active and drag the layer with structural data here. To see both Data Frames (map and projection), it is switched to Layout View." & _
vbNewLine & "Use 'Plot Tool' button. In the dialog box, choose the input layer in the first box. Only layers from the active Data Frame are in the menu. Then choose field for control point ID and fields for S structure and/or L structure. If the table of the input layer doesn't contain fields for control point ID or type of structure, add the field to the table first. Leave it without values. The field for type of structure can be used later by the tool to distinguish between types of plotted structures." & _
vbNewLine & "Untick the option 'Plot only selected features' if you want to plot all (not only selected) features." & _
vbNewLine & "In the Plotting Options, choose the output (graphic OR an existing shapefile  - must have the same table structure as the input layer OR a new shapefile - choose the output workspace and the shapefile name.)" & _
vbNewLine & "Click on 'Plot as points' (orientations of S-structures are plotted as poles to planes) or 'Plot as Arcs' button. (Only planes can be plotted as arcs. Arcs can be plotted only in a shapefile - new or existing one, not as graphics.)"
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        TextBox1.Text = "This tool calculates the density distribution diagram from selected features in the input layer." & _
        vbNewLine & "Create New Projection by the 'New Projection' button first. Make it active and drag the layer with structural data here. To see both Data Frames (map and projection), it is switched to Layout View." & _
        vbNewLine & "Use 'Density distribution' button. In the dialog box, choose the input layer in the first box. Only layers from the active Data Frame are in the menu." & _
        vbNewLine & "Then choose fields for dip direction and dip. Untick the option 'Work with SELECTION' if you want to involve all (not only selected) features into the calculation." & _
        vbNewLine & "Choose the output location and type the name for the output raster. Smoothing parametr: less data - smaller number - more smoothed. It is set in deafult, but try to change, the higher number produces the detailed diagrams. " & _
        vbNewLine & "Click on 'Start' button." & _
        vbNewLine &
        vbNewLine & "The output raster should be exported to save permanently. Add it by the 'Add Data' button to the Projection Data Frame." & _
        vbNewLine & "Appearance Tab: In the Density distrubution dialog box, there is a second tab (page) called 'Appearance'. You can set the classified color ramp to the raster here, to see the contours better."
    End Sub

    Private Sub LinkLabel6_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        TextBox1.Text = "This tool links plotted features in the projection with equivalent features in the map, and vice versa." & _
" Select features in the diagram or map by Selection tools first. Both layers must be in the active Data Frame. Use the 'Link Back' button. In the first box of the dialog box choose the layer with selected features (FROM layer). In the second box choose the layer, you want to link to (TO layer)." & _
" Click on the button representing the link direction ('diagram to map' or 'map to diagram'). Equivalent features are selected."

    End Sub

    Private Sub LinkLabel7_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        TextBox1.Text = "1) Only layers from an ACTIVE Data Frame are offered as input layers in each dialog box. All results are added to an ACTIVE Data Frame." & _
       vbNewLine & "2) Fields for dip and dip direction must by of a double or integer type." & _
       vbNewLine & "3) After you choose the fields names in the boxes, use 'remember button'. Next time use the 'load button' to read the names." & _
       vbNewLine & "4) If an output shapefile name and workspace are identical to existing one, the old shapefile is replaced by the new one." & _
       vbNewLine & "Detailed description is in the OATools guide."
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        TextBox1.Text = "This tool calculates the map of spatially averaged data. " & _
vbNewLine & "It uses selected features from the input layer, if there is no selection, all features are involved. Choose the input layer in the first box, then fields for dip and dip direction." & _
vbNewLine & "The spatial averages are calculated in the regular net of stations. Type the interval of these stations and radius of influence, within the features are involved into the calculation. (Their influence is weighted by the distance from the averaging station.)" & _
        vbNewLine & "Select the output workspace and type the name. There are informative numbers of stations in X and Y directions and the extent of the data in meters in the text boxes." & _
        vbNewLine & "Use 'Calculate' button."
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        TextBox1.Text = "This tool calculates fold axes in the regular net of stations." & _
        vbNewLine & "It uses selected features from the input layer, if there is no selection, all features are involved. Choose the input layer in the first box, then fields for dip and dip direction." & _
        vbNewLine & "The fold axes are calculated in the regular net of stations. Type the interval of these stations and radius of influence, within the features are involved into the calculation. (Their influence is weighted by the distance from the station.)" & _
        vbNewLine & "Select the output workspace and type the name. There are informative numbers of stations in X and Y directions and the extent of the data in meters in the text boxes." & _
        vbNewLine & "Use 'Calculate' button."
    End Sub



    Private Sub LinkLabel5_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        TextBox1.Text = "The 'Rose Histogram' tool enables representation of the frequency of orientations of geological structures in the geological map." & _
        vbNewLine & "Orientations are read from point geometry data (orientations are in the attribute table) or are calculated from the geometry of line data." & _
        vbNewLine & "Dip directions from an attribute table of point data are converted to strikes. In the case of lines, the averaged orientation of line segments is found (the arithmetic mean)." & _
        vbNewLine & "Orientations are matched to one of the 10 degree classes. Class with the highest frequency is 100% of the diagram radius, other classes are relative to it." & _
        vbNewLine & "Select the input layer. Untick 'Only selected', if no selection is made. If input is of a point geometry, select the dip direction field." & _
        vbNewLine & "Select output workspace (folder) and type the name. Use button depending on the geometry of the input data: 'Calculate histogram from lines' or 'Calculate histogram from points'."
    End Sub

    Private Sub HelpForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = "1) Only layers from an ACTIVE Data Frame are offered as input layers in each dialog box. All results are added to an ACTIVE Data Frame." & _
           vbNewLine & "2) Fields for dip and dip direction must by of a double or integer type." & _
           vbNewLine & "3) After you choose the fields names in the boxes, use 'remember button'. Next time use the 'load button' to read the names." & _
           vbNewLine & "4) If an output shapefile name and workspace are identical to existing one, the old shapefile is replaced by the new one." & _
           vbNewLine & "Detailed description is in the OATools guide."
    End Sub

 


    Private Sub LinkLabel9_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel9.LinkClicked
        TextBox1.Text = "There are two additional functions. Averaged orientation and Fold axis. By calculating the averaged orientation, the shape and strength parameters are computed and displayed. The shape parameter higher than 1 means the cluster type of distribution, lower than 1 indicates the girdle type of distribution. The strength parametr indicates the strength of the distribution, the transitional distribution (partly girdle, partly cluster) is around 1. The value near 0 means the uniform/isotropic distribution (Fisher et al., 1987)." & _
vbNewLine & "At least two features must be selected to calculate the fold axis. The calculated fold axis can be plotted in the diagram (as graphic OR in an existing shapefile - must have the same table structure as the input layer OR in a new shapefile). Both the functions use SELECTED features, if the selection is made, or ALL features, if there is no selection in the input layer." & _
vbNewLine & "In this case, the option button 'Plot only selected features' isn´t important, only selected features are always used."
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class