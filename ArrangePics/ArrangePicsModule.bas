Attribute VB_Name = "ArrangePicsModule"
Option Explicit
Option Base 1

Public Sub ArrangePics()

    'Arrange pictures in a smartart
    'Best usage scenario is the add all your customer / product logos and arrange them instantly
    
    'Created by Dr Nitin Paranjape and Raj Chudhuri
    'Created on 8 Jan 2022
    
    'The default Picture SmartArt layout uses Picture Fill.
    'Due to this many pictures get cut off.
    'This macro fits the pictures properly
    
    
    'Declare relevant variables
    
    Dim vw As View
    Dim sld As Slide
    Dim shp As Shape
    
    Dim shpArr() As Shape
    Dim piccount As Integer
    
        
    ' Check if anything is selected
    On Error GoTo ArrangePicsSelectionErr:
    
    If ActiveWindow Is Nothing Then
        MsgBox "Select at least three pictures on a slide and try again"
        Exit Sub
    End If

    If ActiveWindow.Selection Is Nothing Then
        MsgBox "Nothing is selected. Select at least three pictures on a slide and try again"
        Exit Sub
    End If
    
    If ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide Then
        MsgBox "Please switch to normal or slide view." & vbCrLf & _
                "Select at least three pictures on a slide and try again"
        Exit Sub
    End If
        
    On Error GoTo 0
        
    ' Set current view for copy pasting shapes later
    Set vw = ActiveWindow.View
    
    ' Set reference to current slide
    Set sld = ActiveWindow.View.Slide

    
    ' If nothing is selected, return
    If ActiveWindow.Selection.Type = ppSelectionNone Then
        MsgBox "Select multiple pictures and then run this macro"
        Exit Sub
    End If
    
    ' Check each selected item to see if it is a picture or graphic
    ' Only if more than two picture / graphic items are found, the macro works
    
    With ActiveWindow.Selection
        
        If .ShapeRange.Count > 0 Then
            
            piccount = 0
            
            ' Check each shape in the selection
            For Each shp In .ShapeRange
                If isPictureShape(shp) Then
                    ' Increment picture counter
                    piccount = piccount + 1
                    
                    ' Add the shape to array
                    arrayInsert shpArr, shp
                End If
            Next
            
            ' Exit if less than two picture items are available
            If piccount < 3 Then
                MsgBox "Select at least three pictures and try again"
                Exit Sub
            End If
        Else
            MsgBox "Select at least three pictures and try again"
            Exit Sub
        End If
    End With
   
    convertToSmartArt vw, sld, shpArr
    
    Exit Sub
    
ArrangePicsSelectionErr:
    MsgBox "Please switch to normal or slide view." & vbCrLf & _
            "Select at least three pictures on a slide and try again"
End Sub


Private Sub convertToSmartArt(vw As View, sld As Slide, shpArr() As Shape)
    
    ' Add SmartArt of type Bending Picture with Semi Transparent Text
    ' This SmartArt is used because
    '       1. it makes use of available space optimally and
    '       2. the textbox does not occupy extra space. It overlaps the picture.
    ' This layout is identified by the following URN:
    '       urn:microsoft.com/office/officeart/2008/layout/BendingPictureSemiTransparentText
    
    Dim sm As SmartArt
    Dim nd As SmartArtNode
    
    Set sm = sld.Shapes.AddSmartArt(Application.SmartArtLayouts("urn:microsoft.com/office/officeart/2008/layout/BendingPictureSemiTransparentText")).SmartArt
    
    Dim i As Integer
    
    ' Delete all the default blank nodes in the SmartArt
    For i = 1 To sm.AllNodes.Count
        sm.AllNodes(1).Delete
    Next
    
    
    ' Error handler to manage Active Window issues
    On Error GoTo convertToSmartArtTimingErr:
    
    ' Loop through the array of shapes and add the pictures to SmartArt nodes
    For i = 1 To UBound(shpArr)
    
        ' Copy the shape to clipboard
        shpArr(i).Copy
        
        ' Hide the original shape because it is anyway going to be a part of the new SmartArt
        shpArr(i).Visible = False
        
        ' Process events
        DoEvents
        
        ' Add new node to the SmartArt
        Set nd = sm.AllNodes.Add
        
        ' The first shape in the node is text - which is to be kept blank
        ' We use the second shape to add the picture
        
        ' Select the second shape of the node - the picture placeholder
        nd.Shapes(2).Select
        DoEvents
        
        ' Paste the copied picture
        vw.Paste
        DoEvents
        
        ' Make sure that the picture fits completely inside the shape
        ' This will ensure that no part of the picture is cut off and
        ' that the pictures will not be distorted
        CommandBars.ExecuteMso ("PictureFitCrop")
        DoEvents
        
        'Remove border from the picture placeholder
        nd.Shapes(2).Line.Transparency = 1
        
        'Hide the border and fill for the textbox
        nd.Shapes(1).Fill.Transparency = 1
        nd.Shapes(1).Line.Transparency = 1
        
    Next
    
    ' Rename the SmartArt object to "Auto-Diagram"
    sm.Parent.Name = "Auto-Diagram"
    
    ' Select the new SmartArt
    sm.Parent.Select
    
    Exit Sub

convertToSmartArtTimingErr:
    ' This error handler processes some errors which crop up due to executing timing mismatches
    ' We have not found a way to handle these errors in any other way.
    DoEvents
    Resume
End Sub

Private Function isPictureShape(shp As Shape) As Boolean
    Dim result As Boolean
    
    ' Shapes of the following types can contain pictures:
    '   - msoPicture: regular pictures
    '   - msoGraphic: SVG graphics including Icons, Illustrations and Stickers
    '   - msoPlaceholder: placeholders, only if their contained type is picture or graphic
    
    If Not (shp.Type = msoPicture Or shp.Type = msoGraphic Or shp.Type = msoPlaceholder) Then
        result = False
    Else
        result = True
        If shp.Type = msoPlaceholder Then
            If Not ( _
                    shp.PlaceholderFormat.ContainedType = msoPicture Or _
                    shp.PlaceholderFormat.ContainedType = msoGraphic _
                ) Then
                
                result = False
            End If
        End If
    End If
    
    isPictureShape = result
End Function

Private Sub arrayInsert(aShps() As Shape, shp As Shape)
    ' Add selected shapes to the array

    Dim ub As Integer
    
    'If the array not empty
    If Not Not aShps Then
        
        ' Check the new dimension required
        ub = UBound(aShps) + 1
        
        ' Expand the array by 1 more element
        ReDim Preserve aShps(ub)
        
        ' Add the current shape as the last item
        Set aShps(ub) = shp
        
    Else
        ' If array is empty, initialize it with one element
        ReDim aShps(1)
        
        ' Add the current shape as the first element
        Set aShps(1) = shp
        
    End If

End Sub

'End of code
