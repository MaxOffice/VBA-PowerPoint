Attribute VB_Name = "SplitTextModule"
Option Explicit
Option Base 1

Private Const MACROTITLE As String = "Split Text"

Private Enum SplitBy
    Characters
    Words
End Enum

Private Const ErrWrongView = 514
Private Const ErrNoSelection = 515
Private Const ErrSelectOneShape = 516
Private Const ErrNoTextInSelection = 517

Private Const ErrWrongViewText = "Please change to Normal (single slide) View." & vbCrLf & _
                    "Select one shape containing text and try again."
Private Const ErrNoSelectionText = "Currently, no shape or textbox is selected." & vbCrLf & _
                    "Please select one shape containing text and try again."
Private Const ErrSelectOneShapeText = "Please select only one shape or textbox and try again."
Private Const ErrNoTextInSelectionText = "This shape has no text in it." & vbCrLf & _
                    "Please select one shape containing text and try again."



Private Sub SplitText(by As SplitBy)
    'Written by Dr Nitin Paranjape
    '27 Jan 2016
    'use it freely with attribution
    
    'It splits text in a textbox into individual characters
    'For example, if the original textbox contains "demo",
    'SplitText will create four textboxes with "d" "e" "m" "o"
    'This is useful for creating character based animation or text art
    'All the newly created objects are created at 0,0 which is at the top left corner of the current slide
    'The default size is 100,100 (width, height)
    'However, if you have set any default textbox shape, it will be used.
    'Each object gets a name which is the same as the character
    'This name will be visible in the selection pane. Some names will obviously be duplicated.
    'Select all the objects and choose Alignment dropdown
        'first of all select - Align to slide - option
        'now choose distribute horizontally (or vertically) to separate out all the individual shapes
    '------------------------------------------------
    'The code is not tested comprehensively
    'It is designed as a quick fix for a common need
    'Consider it as beta code
    'Use at your own risk
    '------------------------------------------------
    
    'Check if in Normal or Slide view. Else Exit with error
    If Application.ActiveWindow.ViewType <> ppViewNormal Or _
        Application.ActiveWindow.ViewType = ppViewSlide Then
    
        Err.Raise ErrWrongView, MACROTITLE + ".SplitText", ErrWrongViewText
    End If
    
    
    'get selection
    Dim sel As Selection
    Set sel = ActiveWindow.Selection
    
    'check if selection has one shape selected
    If sel.Type = ppSelectionNone Then
        Err.Raise ErrNoSelection, MACROTITLE + ".SplitText", ErrNoSelectionText
    End If
    
    If Not (sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText) Then
        Err.Raise ErrSelectOneShape, MACROTITLE + ".SplitText", ErrSelectOneShapeText
    End If
    
    If sel.ShapeRange.Count <> 1 Then
        Err.Raise ErrSelectOneShape, MACROTITLE + ".SplitText", ErrSelectOneShapeText
    End If
    
    Dim shp As Shape
    Set shp = sel.ShapeRange(1)
    
    If Not shp.HasTextFrame Then
        Err.Raise ErrSelectOneShape, MACROTITLE + ".SplitText", ErrSelectOneShapeText
    End If
    
    
    If Not shp.HasTextFrame Then
        Err.Raise ErrSelectOneShape, MACROTITLE + ".SplitText", ErrSelectOneShapeText
    End If
    
    Dim tr As TextRange2
    Set tr = shp.TextFrame2.textRange
    
    If Len(tr.Text) = 0 Then
        Err.Raise ErrNoTextInSelection, MACROTITLE + ".SplitText", ErrNoTextInSelectionText
    End If
    
    Dim sld As Slide
    Set sld = ActiveWindow.View.Slide
    
    Dim newshp As Shape
    Dim i As Long, newShapeCount As Long
    
    If by = Characters Then
        newShapeCount = tr.Characters.Count
    Else
        newShapeCount = tr.Words.Count
    End If
    
    ReDim arrShapes(newShapeCount) As String
    
    For i = 1 To newShapeCount
         'create new shape
         Set newshp = sld.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 100)
         
         'set the text as the current character or word
         If by = Characters Then
            newshp.TextFrame.textRange.Text = tr.Characters(i).Text
         Else
            newshp.TextFrame.textRange.Text = tr.Words(i).Text
         End If
         
         'add to the shapes array
         arrShapes(i) = newshp.Name
    Next
    
    'select all the newly created shapes
    Dim newShapeRange As ShapeRange
    Set newShapeRange = sld.Shapes.Range(arrShapes)
    
    With newShapeRange
        'distribute these shapes horizontally, across the slide
        .Distribute msoDistributeHorizontally, msoTrue
        .Align msoAlignTops, msoTrue
    End With
    
    'Copy formatting from the base shape
    shp.PickUp
        
    Dim newShape As Shape
    For Each newShape In newShapeRange
        
        With newShape
            'set names for each object for viewing in the Selection Pane
            .Name = .TextFrame.textRange.Text
        
            'apply formatting from original shape
            .Apply
            'this is done to avoid interference with distribute horizontally functionality
            .TextFrame.WordWrap = msoFalse
            .TextFrame.AutoSize = ppAutoSizeShapeToFitText
        End With
    Next
        
    With newShapeRange
        .Select
    
        'distribute these shapes horizontally, across the slide
        .Distribute msoDistributeHorizontally, msoTrue
        .Align msoAlignTops, msoTrue
    End With
    
    'That's it
End Sub


Public Sub SplitText2Words()
    On Error GoTo SplitText2WordsErr
    
    SplitText by:=Words
    
    Exit Sub
SplitText2WordsErr:
    MsgBox Err.Description, vbInformation + vbOKOnly, MACROTITLE
End Sub
    
Public Sub SplitText2Chars()
    On Error GoTo SplitText2CharsErr
    
    SplitText by:=Characters
    
    Exit Sub
SplitText2CharsErr:
    MsgBox Err.Description, vbInformation + vbOKOnly, MACROTITLE
End Sub

