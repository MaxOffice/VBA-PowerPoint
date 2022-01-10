Attribute VB_Name = "ArrangePicsModule"
Option Explicit
Option Base 1


Private Const MAXERRORS As Integer = 100

Public Sub ArrangePics()

    ' Arrange pictures in a smartart
    ' Best usage scenario is the add all your customer / product logos and arrange them instantly
    
    ' Created by Dr Nitin Paranjape and Raj Chaudhuri
    ' Created on 8 Jan 2022
    
    ' How to use this macro:
    ' Select three or more pictures (or graphics) and run the macro
    ' Selected shapes will be arranged in a SmartArt
    ' Original shapes are hidden (not deleted)
    ' There is no technical limit to the number of shapes you can select
    ' More shapes means more processing time.
    
    ' For technical reasons, we cannot show a progress bar.
    ' The only way to check the progress is to look at the status bar
    ' You will see a flickering message like "Press Esc to ..."
    ' This indicates that the processing is in progress
    
    
    ' The default Picture SmartArt layout uses Picture Fill.
    ' Due to this many pictures get cut off.
    ' This macro fits the pictures properly
    
    
    ' Declare relevant variables
    
    Dim vw As View
    Dim sld As Slide
    Dim shp As Shape
    
    ' Array to hold selected shapes
    Dim shpArr() As Shape
    Dim PicCount As Integer
    
        
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
            
            PicCount = 0
            
            ' Check each shape in the selection
            For Each shp In .ShapeRange
                If isPictureShape(shp) Then
                    ' Increment picture counter
                    PicCount = PicCount + 1
                    
                    ' Add the shape to array
                    arrayInsert shpArr, shp
                End If
            Next
            
            ' Exit if less than two picture items are available
            If PicCount < 3 Then
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
    
    Dim errCount As Integer
    
    Set sm = sld.Shapes.AddSmartArt(Application.SmartArtLayouts("urn:microsoft.com/office/officeart/2008/layout/BendingPictureSemiTransparentText")).SmartArt
    
    sm.Parent.Visible = False
    
    Dim i As Integer
    
    ' Delete all the default blank nodes in the SmartArt
    For i = 1 To sm.AllNodes.Count
        sm.AllNodes(1).Delete
    Next
    
    
    ' Error handler to manage Active Window issues
    errCount = 1
    On Error GoTo convertToSmartArtTimingErr:
    
    ' Show progress bar
    Dim pb As ProgressForm
    Set pb = New ProgressForm

    pb.PicCount = UBound(shpArr)
    pb.Show False
    
    ' Loop through the array of shapes and add the pictures to SmartArt nodes
    For i = 1 To UBound(shpArr)
    
        ' Update progress bar
        pb.CurrentPic = i
        
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
        
        ' Remove border from the picture placeholder
        nd.Shapes(2).Line.Transparency = 1
        
        ' Hide the border and fill for the textbox
        nd.Shapes(1).Fill.Transparency = 1
        nd.Shapes(1).Line.Transparency = 1
        
    Next
    
    ' Close progress bar
    Unload pb
    Set pb = Nothing
    

    ' Rename the SmartArt object to "Auto-Diagram"
    sm.Parent.Name = "Auto-Diagram"
    
    ' Select the new SmartArt
    sm.Parent.Visible = True
    
    sm.Parent.Select
    
    Exit Sub

convertToSmartArtTimingErr:
    ' This error handler processes some errors which crop up due to executing timing mismatches.
    ' We have not found a way to handle these errors in any other way.
    ' Therefore, we have implemented a circuit breaker which stops the process after a large
    ' number of errors.
    
    errCount = errCount + 1
    If errCount > MAXERRORS Then
        ' Revert original state
        
        ' Close progress bar if open
        If Not pb Is Nothing Then
            Unload pb
            Set pb = Nothing
        End If
        
        ' Remove the smartart
        Dim sh As Shape
        Set sh = sm.Parent
        sh.Delete
        
        ' Unhide processed pictures
        Dim j As Integer
        For j = 1 To i
            sld.Shapes(j).Visible = msoTrue
        Next
        
        MsgBox "We are sorry. Due to errors beyond our control, the process could not be completed." & vbCrLf & _
                "We have reverted your slide to what it was before the process." & vbCrLf & _
                "The process may work if you try again."
        
        Exit Sub
    End If
    
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


