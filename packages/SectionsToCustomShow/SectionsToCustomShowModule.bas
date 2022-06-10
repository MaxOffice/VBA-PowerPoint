Attribute VB_Name = "SectionsToCustomShowModule"
Option Explicit
Option Base 1

Public Sub SectionsToCustomShows()
' 10 Jun 2022
' Nitin Paranjape
' Creates a custom show for every section.
' If a custom show of the same name exists, it is overwritten

   
    Dim ap As Presentation
    Dim f As Integer
    Dim cs As NamedSlideShow
    Dim css As NamedSlideShows
    Dim secp As SectionProperties
    
    
    Set ap = ActivePresentation
    Set css = ap.SlideShowSettings.NamedSlideShows
    Set secp = ap.SectionProperties
    
    'loop through sections
    'if no sections - error and return
    'for each section
        'create array of slides
        'check if custom show exists
            'if yes, delete it
            'add new section with the slides
    With secp
    
        For f = 1 To .Count
        
            If .SlidesCount(f) > 0 Then
                
                Dim slideCnt As Integer
                slideCnt = .SlidesCount(f)
                    
                Dim sarr() As Long
                Dim g As Integer, cnt As Integer
                
                cnt = 1
                
                ReDim sarr(1 To 1) As Long
                
                
                
                For g = .FirstSlide(f) To .FirstSlide(f) + (.SlidesCount(f) - 1)
                
                    
                    sarr(UBound(sarr)) = ap.Slides(g).SlideID
                    cnt = cnt + 1
                    ReDim Preserve sarr(1 To UBound(sarr) + 1)
                    
                
                Next
                Dim showName As String
                showName = .Name(f)
                
                removeDuplicateShow ap, showName
                
                ap.SlideShowSettings.NamedSlideShows.Add .Name(f), sarr
            
            End If
            
        Debug.Print .SectionID(f)
        Debug.Print .Name(f)
        Debug.Print .FirstSlide(f)
        Debug.Print .SlidesCount(f)
        
        Next
        
        
    End With


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
    Set ap = ActivePresentation
    Dim nmss As NamedSlideShows
    Dim nms As NamedSlideShow
    Set nmss = ap.SlideShowSettings.NamedSlideShows
    
    If nmss.Count < 1 Then
        MsgBox ("No custom shows available to delete")
        Exit Sub
    End If
        
    Stop
    
    If MsgBox("Delete " + Str(nmss.Count) + " custom shows from" + vbCrLf + ap.Name, vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    For Each nms In nmss
    
        
            nms.Delete
        
        
    
    Next
    

End Sub
