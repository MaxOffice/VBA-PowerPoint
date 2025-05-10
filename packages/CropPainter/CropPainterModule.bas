Attribute VB_Name = "CropPainterModule"
Option Explicit

Private Const MACROTITLE = "Crop Painter"

' Variables for holding a crop properties
Private m_ShapeLeft As Single
Private m_ShapeTop As Single
Private m_CropLeft As Single
Private m_CropTop As Single
Private m_CropRight As Single
Private m_CropBottom As Single
Private m_CropPictureOffsetX As Single
Private m_CropPictureOffsetY As Single
Private m_CropPictureWidth As Single
Private m_CropPictureHeight As Single
Private m_CropShapeLeft As Single
Private m_CropShapeTop As Single
Private m_CropShapeWidth As Single
Private m_CropShapeHeight As Single
Private m_CropCopied As Boolean


Public Sub CropPaintCopy()
    On Error GoTo CropCopyErr
    
    Dim sourceShape As shape
    Set sourceShape = SelectedPictureShape
    
    If sourceShape Is Nothing Then
        MsgBox "Select a single picture on a slide and try again.", vbExclamation, MACROTITLE
        Exit Sub
    End If
   
    m_ShapeLeft = sourceShape.Left
    m_ShapeTop = sourceShape.Top
    m_CropLeft = sourceShape.PictureFormat.CropLeft
    m_CropTop = sourceShape.PictureFormat.CropTop
    m_CropRight = sourceShape.PictureFormat.CropRight
    m_CropBottom = sourceShape.PictureFormat.CropBottom
    m_CropPictureOffsetX = sourceShape.PictureFormat.Crop.PictureOffsetX
    m_CropPictureOffsetY = sourceShape.PictureFormat.Crop.PictureOffsetY
    m_CropPictureWidth = sourceShape.PictureFormat.Crop.PictureWidth
    m_CropPictureHeight = sourceShape.PictureFormat.Crop.PictureHeight
    m_CropShapeLeft = sourceShape.PictureFormat.Crop.ShapeLeft
    m_CropShapeTop = sourceShape.PictureFormat.Crop.ShapeTop
    m_CropShapeWidth = sourceShape.PictureFormat.Crop.ShapeWidth
    m_CropShapeHeight = sourceShape.PictureFormat.Crop.ShapeHeight
    m_CropCopied = True

    Exit Sub
CropCopyErr:
    MsgBox "An error occured while trying to copy Crop properties: " & Err.Description, _
            vbExclamation + vbOKOnly, MACROTITLE
End Sub

Public Sub CropPaintPaste()
    On Error GoTo CropPasteErr
    
    If Not m_CropCopied Then
        MsgBox "First, select the source picture and copy crop properties.", _
                vbInformation + vbOKOnly, MACROTITLE
        Exit Sub
    End If
    
    Dim selectedRange As ShapeRange
    Set selectedRange = SelectedShapeRange
    
    If selectedRange Is Nothing Then
        MsgBox "Select one or more pictures, and then try again.", _
                    vbInformation + vbOKOnly, MACROTITLE
        Exit Sub
    End If
    
    Dim shape As shape
    For Each shape In selectedRange
        If shape.Type = msoPicture Then
            ' First, "paste" shape position
            shape.Left = m_ShapeLeft
            shape.Top = m_ShapeTop
            With shape.PictureFormat
                ' Then, the crop position
                .CropLeft = m_CropLeft
                .CropTop = m_CropTop
                .CropRight = m_CropRight
                .CropBottom = m_CropBottom
                ' Then, the crop shape
                .Crop.ShapeLeft = m_CropShapeLeft
                .Crop.ShapeTop = m_CropShapeTop
                .Crop.ShapeWidth = m_CropShapeWidth
                .Crop.ShapeHeight = m_CropShapeHeight
                ' Then, the adjustment of the picture
                ' in the crop shape
                .Crop.PictureWidth = m_CropPictureWidth
                .Crop.PictureHeight = m_CropPictureHeight
                .Crop.PictureOffsetX = m_CropPictureOffsetX
                .Crop.PictureOffsetY = m_CropPictureOffsetY
                ' The order is important
            End With
        End If
    Next
    Exit Sub
CropPasteErr:
    MsgBox "An error occured while trying to paste Crop properties: " & Err.Description, _
            vbExclamation + vbOKOnly, MACROTITLE
End Sub

Public Sub CropPaintReset()
    m_CropCopied = False
End Sub

Private Function SelectedShapeRange() As ShapeRange
    Set SelectedShapeRange = Nothing
    
    If ActiveWindow Is Nothing Then
        Exit Function
    End If

    If ActiveWindow.Selection Is Nothing Then
        Exit Function
    End If

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Exit Function
    End If
    
    Set SelectedShapeRange = ActiveWindow.Selection.ShapeRange
End Function

Private Function SelectedPictureShape() As shape
    Set SelectedPictureShape = Nothing
    
    Dim selectedRange As ShapeRange
    Set selectedRange = SelectedShapeRange
    
    If selectedRange Is Nothing Then
        Exit Function
    End If
    
    If selectedRange.Count <> 1 Then
        Exit Function
    End If
    
    Dim selectedShape As shape
    Set selectedShape = selectedRange(1)
    
    If selectedShape.Type <> msoPicture Then
        Exit Function
    End If
    
    Set SelectedPictureShape = selectedShape
End Function

