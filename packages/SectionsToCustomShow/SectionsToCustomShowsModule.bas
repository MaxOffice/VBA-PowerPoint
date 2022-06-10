Attribute VB_Name = "SectionsToCustomShowsModule"
Option Explicit
Option Base 1

Private Const MACROTITLE As String = "Sections To Custom Shows"

Public Sub SectionsToCustomShows()
    ' 10 Jun 2022
    ' Nitin Paranjape
    ' Creates a custom show for every section.
    ' If a custom show of the same name exists, it is overwritten

   
    Dim ap As Presentation
    Dim nss As NamedSlideShows
    Dim secp As SectionProperties
    
    ' Error handler in case this is run from an Add-in, and
    ' no presentation is currently selected
    On Error GoTo SectionsToCustomShowsSelectionErr
    
    Set ap = ActivePresentation
    Set nss = ap.SlideShowSettings.NamedSlideShows
    Set secp = ap.SectionProperties
    
    ' At this point, the presence of an active presentation is established
    ' So we turn the selection error handler off. This way, we can get to
    ' know about errors we have not thought of.
    On Error GoTo 0
    
    ' loop through sections
    ' if no sections - error and return
    ' for each section
        ' create array of slides
        ' check if custom show exists
            ' if yes, delete it
            ' add new section with the slides
    Dim sectionCount As Long
    Dim sectionSlideCount As Long
    Dim sectionFirstSlide As Long
    Dim sectionLastSlide As Long
    Dim showName As String
    Dim j As Long
    Dim i As Long
    Dim convertCount As Long
    
    convertCount = 0
    
    With secp
        sectionCount = .Count
        
        If sectionCount = 0 Then
            MsgBox "There are no sections in this presentation.", _
                    vbInformation + vbOKOnly, _
                    MACROTITLE
                    
            Exit Sub
        End If
        
        For i = 1 To sectionCount
            sectionSlideCount = .SlidesCount(i)
            
            If sectionSlideCount > 0 Then
                sectionFirstSlide = .FirstSlide(i)
                sectionLastSlide = sectionFirstSlide + (sectionSlideCount - 1)
                showName = .Name(i)

                ReDim sarr(sectionFirstSlide To sectionLastSlide)
                
                For j = sectionFirstSlide To sectionLastSlide
                    sarr(j) = ap.Slides(j).SlideID
                Next
                
                removeDuplicateShow ap, showName
                
                nss.Add showName, sarr
                convertCount = convertCount + 1
            End If
        
        Next
    End With

    If convertCount > 0 Then
        MsgBox "Success! Created " & convertCount & " custom shows from sections.", _
                vbInformation + vbOKOnly, _
                MACROTITLE
    End If
    
    Exit Sub
SectionsToCustomShowsSelectionErr:
    MsgBox "Please select a slide in the normal view in any presentation, and try again.", _
            vbExclamation, _
            MACROTITLE
End Sub


Private Sub removeDuplicateShow(subap As Presentation, nm As String)
    Dim nms As NamedSlideShow
    
    For Each nms In subap.SlideShowSettings.NamedSlideShows
        If UCase(nm) = UCase(nms.Name) Then
            nms.Delete
        End If
    Next
End Sub


Public Sub DeleteAllCustomShows()
    Dim ap As Presentation
    Dim nmss As NamedSlideShows
    Dim nms As NamedSlideShow
    Dim i As Long
    Dim showCount As Long
    
    ' Error handler in case this is run from an Add-in, and
    ' no presentation is currently selected
    On Error GoTo DeleteAllCustomShowsSelectionErr
    Set ap = ActivePresentation
    
    
    Set nmss = ap.SlideShowSettings.NamedSlideShows
    showCount = nmss.Count
    
    ' At this point, the presence of an active presentation is established
    ' So we turn the selection error handler off. This way, we can get to
    ' know about errors we have not thought of.
    On Error GoTo 0
        
    If showCount < 1 Then
        MsgBox "No custom shows available to delete.", _
                vbInformation + vbOKOnly, _
                MACROTITLE
        
        Exit Sub
    End If
   
    If MsgBox( _
            "Delete " & showCount & " custom shows from " & vbCrLf & ap.Name & "?", _
            vbQuestion + vbYesNo, _
            MACROTITLE _
        ) <> vbYes Then
        
        Exit Sub
    End If
    
    
    For i = 1 To showCount
        nmss(1).Delete
    Next
    
    Exit Sub
    
DeleteAllCustomShowsSelectionErr:
    MsgBox "Please select a slide in the normal view in any presentation, and try again.", _
            vbExclamation, _
            MACROTITLE
End Sub
