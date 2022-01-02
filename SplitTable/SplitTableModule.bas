Attribute VB_Name = "SplitTableModule"
Option Explicit

Public Sub SplitTable()
    ' If current window and selection cannot be determined, do
    ' nothing and exit.
    If ActiveWindow Is Nothing Then
        Exit Sub
    End If

    If ActiveWindow.Selection Is Nothing Then
        Exit Sub
    End If
    
    ' Check if the selection is exactly one table shape. If not,
    ' show message and exit. Else, process it.
    With ActiveWindow.Selection
        If .Type <> ppSelectionShapes Then
            MsgBox "Select only and exactly one table, and try again", vbExclamation, "Split Table"
            Exit Sub
        End If
    
        If .ShapeRange.Count <> 1 Then
            MsgBox "Select only and exactly one table, and try again", vbExclamation, "Split Table"
            Exit Sub
        End If
        
        Dim selectedShape As Shape
        Set selectedShape = .ShapeRange(1)
        If selectedShape.HasTable <> msoTrue Then
            MsgBox "The selection does not seem to be a table." & vbCrLf & _
                   "Select only and exactly one table, and try again", vbExclamation, "Split Table"
            Exit Sub
        End If
        
        explodeTable selectedShape.Table
        
        selectedShape.Visible = msoFalse
    End With
End Sub

Private Sub explodeTable(ByVal tbl As Table)
    Dim rows As Integer, cols As Integer
    Dim i As Integer, j As Integer

    rows = tbl.rows.Count
    cols = tbl.Columns.Count

    For i = 1 To rows
        For j = 1 To cols
            duplicateCell tbl, i, j
        Next
    Next
End Sub

Private Sub duplicateCell(tbl As Table, curRow As Integer, curCol As Integer)
    Dim newTable As Table
    
    ' Make a copy of the original table
    Set newTable = tbl.Parent.Duplicate(1).Table
    
    With newTable
        Dim rows As Integer, cols As Integer
        Dim i As Integer, j As Integer
    
        rows = .rows.Count
        cols = .Columns.Count
        
        ' Delete rows before and after the current one
        For i = 1 To curRow - 1
            .rows(1).Delete
        Next
        
        For i = curRow + 1 To rows
            .rows(2).Delete
        Next
        
        'Delete columns before and after the current one
       If .Columns.Count > 1 Then
       
       
            For j = 1 To curCol - 1
                
                .Columns(1).Delete
                
                If .Columns.Count = 1 Then
                    Exit For
                End If
            Next
        
        
        
            If .Columns.Count > 1 Then
                
                For j = curCol + 1 To cols
                    
                    .Columns(2).Delete
                    
                    If .Columns.Count = 1 Then
                        Exit For
                    End If
                Next
            
            End If
            
        End If '.columns.count > 1 - first one
        
        ' Change dimensions of the single-cell table shape to
        ' match the dimensions of the correspoding cell in
        ' the original table.
        Dim originalShape As Shape
        Set originalShape = tbl.Cell(curRow, curCol).Shape
        'originalShape.PickUp
        With .Parent
            .Left = originalShape.Left
            .Top = originalShape.Top
            .Width = originalShape.Width
            .Height = originalShape.Height
            
            .Name = tbl.Parent.Name + " >> R:" + Trim(Str(curRow)) + " C:" + Trim(Str(curCol))
            
        End With
        
        
    End With
    
End Sub

