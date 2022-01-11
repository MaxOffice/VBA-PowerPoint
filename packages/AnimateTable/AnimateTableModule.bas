Attribute VB_Name = "AnimateTableModule"
Option Explicit

Private Const MACROTITLE As String = "Animate Table"

Public Sub AnimateTable()
    
    'Split a table into separate objects (tables) for each cell and
    'animate each cell
    
    'How to use this macro?
    '- Select a single table and run this macro
    '- The original table will be hidden. You can unhide it using the Selection Pane.
    '- Individual cells will be created as separate objects
    '- Each cell has Appear animation applied automatically
    '- You can customize the animation as required or remove the animation completely
    
    
    'Performance
    'If the table has many rows and columns, the processing will take some time.
    
    'Extensibility
    'This macro copies only the background color or gradient, font, font color and borders.
    'If you want to copy some other attributes, add more code to the "copyShapeFormatting" subroutine
    
    
    'Created by Dr Nitin Paranjape and Raj Chaudhuri
    'Created on Jan 4, 2022
    'If current window and selection cannot be determined,
    'do nothing and exit.
    
    
    'check if anything is selected
    If ActiveWindow Is Nothing Then
        Exit Sub
    End If

    If ActiveWindow.Selection Is Nothing Then
        Exit Sub
    End If
    
    If ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide Then
        MsgBox "Please switch to normal or slide view." & vbCrLf & _
                "Select only and exactly one table, and try again.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    ' Check if the slide pane is active (not the thumnail or any other pane)
    If Not ActiveWindow.ActivePane.ViewType = ppViewSlide Then
        MsgBox "Select only and exactly one table, and try again.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    ' Check if the selection is exactly one table shape.
    ' If not, show message and exit. Else, process it.
    
    With ActiveWindow.Selection
    
        'table is a shape. But if table is being edited, selection is text
        If .Type <> ppSelectionShapes And .Type <> ppSelectionText Then
            MsgBox "Select only and exactly one table, and try again", vbExclamation, MACROTITLE
            Exit Sub
        End If

        'if multiple shapes are selected, exit
        If .ShapeRange.Count <> 1 Then
            MsgBox "Select only and exactly one table, and try again", vbExclamation, MACROTITLE
            Exit Sub
        End If

        
        'Create shape object for further processing
        Dim selectedShape As Shape
        Set selectedShape = .ShapeRange(1)
        
        If selectedShape.HasTable <> msoTrue Then
            MsgBox "The selection does not seem to be a table." & vbCrLf & _
                   "Select only and exactly one table, and try again", vbExclamation, MACROTITLE
            Exit Sub
        End If
        
        
        'If table has only one row and column there is no need to split
        With selectedShape.Table
            If .rows.Count = 1 And .Columns.Count = 1 Then
                MsgBox "This table has only one row and column." & vbCrLf & _
                        "Cannot split this table further." & vbCrLf & _
                        "Just add animation to the table as desired.", _
                        vbInformation, MACROTITLE
                Exit Sub
            End If
            
        End With
        
        'Process and explode the table
        explodeTable selectedShape.Table
        
        'Hide the original table
        selectedShape.Visible = msoFalse
    
    End With

End Sub

Private Sub explodeTable(ByVal tbl As Table)
    Dim rows As Integer, cols As Integer
    Dim i As Integer, j As Integer
    
    'Intermediate references to understanding processing progress
    Dim topOffset As Single, leftOffset As Single
    Dim prevMergedColumnCount As Integer, totalMergedColumnCount As Integer

    rows = tbl.rows.Count
    cols = tbl.Columns.Count
    
    topOffset = tbl.Parent.Top
    
    'Process each row
    For i = 1 To rows
        leftOffset = tbl.Parent.Left
        prevMergedColumnCount = 0
        totalMergedColumnCount = 0
            
        'Process each column
        For j = 1 To cols
        
            ' Check if the current cell is a merged cell,
            ' by comparing its top and left positions to
            ' what would have been the position in an
            ' evenly spaced table. This is a risk, but
            ' the PowerPoint object model does not seem
            ' to have a better method.
            ' Since all dimensions in VBA are represented
            ' as single-precision numbers, we need to
            ' round them off to avoid problems with very
            ' small variations.
            Dim currCellShape As Shape
            Set currCellShape = tbl.Cell(i, j).Shape

            If Round(currCellShape.Top, 3) < Round(topOffset, 3) Then
                ' If the top position of the current cell
                ' is less than the maximum top position
                ' processed so far, then the current cell
                ' is likely to be merged across rows
                Debug.Print "Merged row detected. Cell R" & i & "C" & j & " skipped."
            ElseIf Round(currCellShape.Left, 3) < Round(leftOffset, 3) Then
                ' If the left position of the current cell
                ' is less than the maximum left position
                ' processed so far, then the current cell
                ' is likely to be merged across columns
                Debug.Print "Merged column detected. Cell R" & i & "C" & j & " skipped."
                prevMergedColumnCount = prevMergedColumnCount + 1
                totalMergedColumnCount = totalMergedColumnCount + 1
            Else
                'process each cell
                duplicateCell tbl, i, j, prevMergedColumnCount, totalMergedColumnCount
                
                'adjust left offset
                leftOffset = leftOffset + currCellShape.Width
                prevMergedColumnCount = 0
            End If
        Next
        
        'Adjust the top offset
        topOffset = topOffset + tbl.rows(i).Height
    Next
End Sub

Private Sub duplicateCell(tbl As Table, curRow As Integer, curCol As Integer, prevMergedColCount As Integer, totalMergedColCount As Integer)
    
    'Overall logic
        'Duplicate the table.
        'Remove unwanted rows and columns
        'Retain the currently desired cell
        'Position it exactly where the original cell is
        'Copy formatting from original cell
        'Apply Appear animation effect
    
    
    Dim newTable As Table
    
    ' Make a copy of the original table
    Set newTable = tbl.Parent.Duplicate(1).Table
    
    With newTable
    
        'Remove unwanted formatting styles from the duplicate table
        'This is necessary because when a row or column is deleted, default formatting shifts.
        
        .FirstRow = False
        .LastRow = False
        .HorizBanding = False
        .VertBanding = False
        .LastCol = False
        .FirstCol = False
        
        
        
        
        Dim rows As Integer, cols As Integer
        Dim i As Integer, j As Integer
    
        rows = .rows.Count
        cols = .Columns.Count
        
        ' Delete rows before the current one
        For i = 1 To curRow - 1
            .rows(1).Delete
        Next
        
        ' If there are merged columns in the current row,
        ' or split columns before the current row, then
        ' deleting the rows before it may cause the column
        ' count to change. If this is the case, then we
        ' should pick up the new column count, and use
        ' that instead of the original.
        
        ' If the original column count at this point is
        ' more than the table column count, the current
        ' row contains at least one merged cell, or
        ' previous rows had split cells.
        If cols > .Columns.Count Then
            cols = .Columns.Count
            
            ' If the current column, less total merged
            ' columns, is greater than the new column
            ' count, we should delete the duplicate
            ' table itself, and not proceed.
            If (curCol - totalMergedColCount) > cols Then
                .Parent.Delete
                Exit Sub
            End If
        End If
        
        For i = curRow + 1 To rows
            .rows(2).Delete
        Next
        
        ' If there are split cells after the current row,
        ' deleting the rows after it may cause the column
        ' count to change again.
        If cols > .Columns.Count Then
            cols = .Columns.Count
        End If
        
        ' Delete columns before the current one,
        ' excluding merged columns.
        For j = 1 To curCol - (prevMergedColCount + 1)
            .Columns(1).Delete
            ' Because of the possibility of merged
            ' cells, doing this can reduce the
            ' column count to 1, in which case we
            ' should stop.
            If .Columns.Count = 1 Then
                Exit For
            End If
        Next
     
        ' If there are any columns left after the
        ' current one, delete them
        If .Columns.Count > 1 Then
            Dim startCol As Integer
            
            ' If there were any previously merged cells
            ' before the current one, the column numbers
            ' need to be adjusted.
            If prevMergedColCount > 0 Then
                startCol = curCol - prevMergedColCount
            Else
                startCol = curCol + 1
            End If
            
            For j = startCol To cols
                .Columns(2).Delete
                ' Because of the possibility of merged
                ' cells, doing this can reduce the
                ' column count to 1, in which case we
                ' should stop.
                If .Columns.Count = 1 Then
                    Exit For
                End If
            Next
        End If
        
        'Copy formatting from original to new shape
        Dim originalShape As Shape, newShape As Shape
        
        Set originalShape = tbl.Cell(curRow, curCol).Shape
        
        Set newShape = newTable.Cell(1, 1).Shape
        
        copyShapeFormatting originalShape, newShape
        
        ' Change dimensions of the single-cell table shape to
        ' match the dimensions of the correspoding cell in
        ' the original table.
        With .Parent
            .Left = originalShape.Left
            
            .Top = originalShape.Top
            .Width = originalShape.Width
            .Height = originalShape.Height
            
            .Name = tbl.Parent.Name + " >> R:" + Trim(Str(curRow)) + " C:" + Trim(Str(curCol))
            
        End With
        
        'animate the new table piece
        animateShape newTable
        
    End With
    
End Sub

Private Sub copyShapeFormatting(origShape As Shape, newShape As Shape)
    'Copy Font, Fill and Border formatting from original table cell

    'Suppresses errors due to non-existent properties like border: shape.line
    
    On Error Resume Next
    
    
    
    With newShape
    
        'Copy font formatting
        .TextFrame.TextRange.Font.Color = origShape.TextFrame.TextRange.Font.Color
        .TextFrame.TextRange.Font.Name = origShape.TextFrame.TextRange.Font.Name
        
        'Copy border formatting
        .Line.ForeColor.RGB = origShape.Line.ForeColor.RGB
        .Line.BackColor.RGB = origShape.Line.BackColor.RGB
        .Line.DashStyle = origShape.Line.DashStyle
        
        'Copy background color / gradient
        .Fill.ForeColor.RGB = origShape.Fill.ForeColor.RGB
        .Fill.BackColor.RGB = origShape.Fill.BackColor.RGB
   
    End With
    
End Sub

Private Sub animateShape(sh As Table)
    
    'Animate the shape with default - appear animation
    Dim sld As Slide
    Dim eff As Effect
    Set sld = sh.Parent.Parent
    
    Set eff = sld.TimeLine.MainSequence.AddEffect(sh.Parent, msoAnimEffectAppear)
    
End Sub
