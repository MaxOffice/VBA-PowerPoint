Attribute VB_Name = "RemoveUnusedLayoutsModule"
Option Explicit

Private Const MACROTITLE As String = "Remove Unused Layouts"


Public Sub RemoveUnusedLayouts()
    Dim ap As Presentation
    Dim sld As Slide
    Dim master As Design
    Dim layout As CustomLayout
    Dim uniqueID As String
    Dim usedLayoutIndices As Object
    Dim i As Long, j As Long
    Dim totalMasters As Long, totalLayouts As Long
    Dim unusedCount As Long
    
    On Error GoTo RemoveUnusedLayoutsErr
    
    Set ap = ActivePresentation
    If ap Is Nothing Then
        MsgBox "Please open a presentation, and try again.", vbOKOnly, MACROTITLE
        Exit Sub
    End If

    Set usedLayoutIndices = CreateObject("Scripting.Dictionary")

    ' Identify all used layout indices
    For Each sld In ap.Slides
        ' Use a combination of the layout's index and the design's index as a unique id
        uniqueID = sld.CustomLayout.Index & "_" & sld.CustomLayout.Design.Index
        If Not usedLayoutIndices.Exists(uniqueID) Then
            usedLayoutIndices.Add uniqueID, True ' sld.CustomLayout.Index, sld.CustomLayout.Design.Index
        End If
    Next sld
    
    ' Count used/vs unused
    
    totalMasters = ap.Designs.Count
    totalLayouts = 0
    For Each master In ap.Designs
        totalLayouts = totalLayouts + master.SlideMaster.CustomLayouts.Count
    Next
    unusedCount = totalLayouts - usedLayoutIndices.Count
    
    ' Ask for confirmation
    Dim message As String
    message = "There are " & totalMasters & " designs " & "with " & totalLayouts & " layouts in this presentation." & vbCrLf & _
              "Out of these, " & unusedCount & " are unused." & vbCrLf & _
              "Should I delete them ?"
    
    If MsgBox(message, vbYesNo + vbDefaultButton1, MACROTITLE) <> vbYes Then
        Exit Sub
    End If

    ' Attempt to delete unused layouts
    For Each master In ap.Designs
        For i = master.SlideMaster.CustomLayouts.Count To 1 Step -1
            ' Use a combination of the layout's index and the design's index as a unique id
            uniqueID = i & "_" & master.Index
            If Not usedLayoutIndices.Exists(uniqueID) Then
                ' Attempt to delete the layout
                On Error Resume Next ' In case of error, skip to the next layout
                master.SlideMaster.CustomLayouts(i).Delete
                If Err.Number <> 0 Then
                    Debug.Print "Could not delete layout at index: " & i
                    Err.Clear
                End If
                On Error GoTo RemoveUnusedLayoutsErr:
            End If
        Next i
    Next master

    MsgBox "Attempted to delete unused layouts. Please review your presentation." & vbCrLf & _
           "There should be " & usedLayoutIndices.Count & " layouts in " & ap.Designs.Count & " designs left." & vbCrLf & _
           "If you do not like the results, you can still perform an Undo operation.", vbOKOnly, MACROTITLE
    Exit Sub
    
RemoveUnusedLayoutsErr:
    MsgBox "Error while performing operation: " & Err.Description & vbCrLf & _
           "Please try again.", vbOKOnly + vbExclamation, MACROTITLE
End Sub

