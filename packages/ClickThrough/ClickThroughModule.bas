Attribute VB_Name = "ClickThroughModule"
Option Explicit
Option Base 1



Private Const MACROTITLE As String = "Click-through Maker"

Public Sub AnimatePics()

    ' Animate selected pictures to show one by one on click
    
    ' Easy way to create a click-through or flipbook type of slide
    
    ' Created by Dr Nitin Paranjape and Raj Chaudhuri
    ' Created on 14 Jan 2022
    
    ' How to use this macro:
    ' Select two or more pictures (or graphics) and run the macro
    ' Items which do not contain pictures are ignored
    ' Each picture gets
    '   entry / exit animation
    '   Drop Shadow rectangle style
    
    ' Existing animations, if any, are not removed.
    
    ' Pictures are arranged in the order in which they were added to the slide
    ' To animate in a custom order, manually select pictures in that order one by one
    ' Click the first picture and then shift click to select subsequent pictures
    
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
        MsgBox "Select at least two pictures on a slide and try again.", vbExclamation, MACROTITLE
        Exit Sub
    End If

    If ActiveWindow.Selection Is Nothing Then
        MsgBox "Nothing is selected. Select at least two pictures on a slide and try again." _
                    , vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    If ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide Then
        MsgBox "Please switch to normal or slide view." & vbCrLf & _
                "Select at least two pictures on a slide and try again.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    ' Check if the slide pane is active (not the thumnail or any other pane)
    If Not ActiveWindow.ActivePane.ViewType = ppViewSlide Then
        MsgBox "Select at least two pictures on a slide and try again.", vbExclamation, MACROTITLE
        Exit Sub
    End If
        
    On Error GoTo 0
        
    ' Set current view for copy pasting shapes later
    Set vw = ActiveWindow.View
    
    ' Set reference to current slide
    Set sld = ActiveWindow.View.Slide
    
    
    
    ' If nothing is selected, return
    If ActiveWindow.Selection.Type = ppSelectionNone Then
        MsgBox "Select at least two pictures on a slide and try again.", vbExclamation, MACROTITLE
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
            If PicCount < 2 Then
                MsgBox "Select at least two pictures on a slide and try again.", vbExclamation, MACROTITLE
                Exit Sub
            End If
        Else
            MsgBox "Select at least two pictures on a slide and try again.", vbExclamation, MACROTITLE
            Exit Sub
        End If
    End With
   
    animPicMain vw, sld, shpArr
    
    Exit Sub
    
ArrangePicsSelectionErr:
    MsgBox "Please switch to normal or slide view." & vbCrLf & _
            "Select at least two pictures on a slide and try again.", vbExclamation, MACROTITLE
End Sub


Private Sub animPicMain(vw As View, sld As Slide, aShps() As Shape)
    
    ' Animate pictures available in the array aShps
    
    Dim xe As Effect
    Dim ff As Integer
    
    
    ' Show progress bar (non modal)
    Dim pb As ProgressForm
    Set pb = New ProgressForm
    pb.Caption = MACROTITLE
    pb.PicCount = UBound(aShps)
    pb.Show False
    
    ' Iterate images array
    ' Add effect and animation
    
    For ff = 1 To UBound(aShps)
    
        ' Update progress bar
        pb.CurrentPic = ff
        
        ' Apply shadow effect
        applyShadow aShps(ff)
        
        ' No animation needed for the first image
        If ff > 1 Then ' 2nd to last - add entry effect
            
            ' Add entry to current item
            Set xe = sld.TimeLine.MainSequence.AddEffect(aShps(ff), msoAnimEffectAppear)
            
            'add exit to previous item
            
                Set xe = sld.TimeLine.MainSequence.AddEffect(aShps(ff - 1), msoAnimEffectAppear, , msoAnimTriggerWithPrevious)
                xe.Exit = True
            
            
        End If 'ff>1
    
        DoEvents
        
    Next 'ff = 1 to cnt
    
    ' Close progress bar
    Unload pb
    Set pb = Nothing
    
    '<--------------
    Exit Sub

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

Private Function applyShadow(pic As Shape)

    ' Applies shadow like the picture style "Drop Shadow Rectangle"
    
    With pic.Shadow
        
        .Type = msoShadow25
        .Blur = 15
        .Transparency = 0.3
        .Size = 100
           
    
    End With

End Function


'End of code


