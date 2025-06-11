Attribute VB_Name = "SelSameLayoutModule"
Option Explicit

Private Const MACROTITLE As String = "Select Same Layout"

Public Sub SelectSlidesWithSameLayout()
    Dim currentLayout As CustomLayout
    Dim sld As Slide
    Dim matchingSlides As Collection
    Dim sldRange As SlideRange
    Dim i As Integer

    On Error GoTo SelectSlidesWithSameLayoutErr
    
    ' Ensure a slide is selected
    If ActiveWindow.Selection.Type <> ppSelectionSlides Then
        MsgBox "Please select a slide first.", _
                vbExclamation + vbOKOnly, MACROTITLE
        Exit Sub
    End If

    ' Get the CustomLayout of the currently selected slide
    Set currentLayout = ActiveWindow.Selection.SlideRange(1).CustomLayout

    ' Collect slides with the same layout
    Set matchingSlides = New Collection
    For Each sld In ActivePresentation.Slides
        If sld.CustomLayout Is currentLayout Then
            matchingSlides.Add sld
        End If
    Next

    ' Build a SlideRange from the matching slides
    If matchingSlides.Count > 1 Then
        Dim slideIndexes() As Integer
        ReDim slideIndexes(1 To matchingSlides.Count)
        For i = 1 To matchingSlides.Count
            slideIndexes(i) = matchingSlides(i).SlideIndex
        Next
        Set sldRange = ActivePresentation.Slides.Range(slideIndexes)
        sldRange.Select
        MsgBox matchingSlides.Count & " slides selected.", _
                vbInformation + vbOKOnly, MACROTITLE
    Else
        MsgBox "No other slides found with the same layout.", _
                vbInformation + vbOKOnly, MACROTITLE
    End If
    
    Exit Sub
SelectSlidesWithSameLayoutErr:
    MsgBox "An error occured: " & Err.Description, _
            vbExclamation + vbOKOnly, MACROTITLE
End Sub

