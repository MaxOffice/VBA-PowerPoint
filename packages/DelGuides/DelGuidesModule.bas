Attribute VB_Name = "DelGuidesModule"
Option Explicit
Option Base 1

Private Const MACROTITLE As String = "Delete All Guides"

Public Sub DelGuides()
    'deletes all guides

    Dim ap As Presentation
    Dim wn As SlideShowWindow
    Dim sld As Slide
    Dim shp, shp2 As Shape
    Dim x, ff, cnt As Integer
    Dim gd As Guide
    Dim dsg As Design
    Dim slm As master
    Dim cl As CustomLayout
    
    On Error GoTo DelGuidesErr
    Set ap = ActivePresentation
    If ap Is Nothing Then
        MsgBox "Please open a presentation, and try again.", vbOKOnly, MACROTITLE
        Exit Sub
    End If
      
        
    If MsgBox( _
                "Delete ALL guides from " & ap.Name & "?" & vbCrLf & _
                "Please note that this cannot be undone.", _
                vbQuestion + vbYesNo, _
                MACROTITLE _
            ) <> vbYes Then
            
        Exit Sub
    End If
        
    For Each dsg In ap.Designs
        Set slm = dsg.slideMaster
        
        'remove from all slide masters
        DeleteGuides slm.Guides
        
        For Each cl In slm.CustomLayouts
            'remove from all customlayouts
            DeleteGuides cl.Guides
        Next cl
    
    Next dsg
    
    'now remove at presentation level
    DeleteGuides ap.Guides
    
    Exit Sub
DelGuidesErr:
    MsgBox "Error while performing operation: " & Err.Description & vbCrLf & _
           "Please try again.", vbOKOnly + vbExclamation, MACROTITLE
End Sub

Private Sub DeleteGuides(g As Guides)
    Dim slideCount, i As Long
    
    slideCount = g.Count
    If slideCount = 0 Then
        Exit Sub
    End If
    
    For i = 1 To slideCount
        g(1).Delete
    Next
End Sub
