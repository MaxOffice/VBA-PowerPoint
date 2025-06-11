Attribute VB_Name = "SectionsFromBulletsModule"
Option Explicit

Private Const MACROTITLE = "Sections From Bullets"

Public Sub CreateSectionsFromBullets()
    On Error GoTo CreateSectionsFromBulletsErr
    
    If ActiveWindow Is Nothing Then
        MsgBox "Please select a slide which has a bulleted list, and try again.", _
                    vbInformation + vbOKOnly, MACROTITLE
        Exit Sub
    End If
    
    Dim newsectioncount As Long
    newsectioncount = ProcessSlideIntoSections(ActiveWindow.View.Slide, processLowerLevels:=False, insertSlidesOnly:=False)
    If newsectioncount = 0 Then
        MsgBox "No top-level bullet points found in the current slide", vbInformation + vbOKOnly, MACROTITLE
    End If
    
    Exit Sub
CreateSectionsFromBulletsErr:
    MsgBox "Please switch to normal or slide view in any presentation." & vbCrLf & _
        "Select a slide which has a bulleted list and try again.", vbExclamation, MACROTITLE
End Sub

Public Sub CreateSlidesFromBullets()
    On Error GoTo CreateSlidesFromBulletsErr
    
    If ActiveWindow Is Nothing Then
        MsgBox "Please select a slide which has a bulleted list, and try again.", _
                    vbInformation + vbOKOnly, MACROTITLE
        Exit Sub
    End If
    
    Dim newsectioncount As Long
    newsectioncount = ProcessSlideIntoSections(ActiveWindow.View.Slide, processLowerLevels:=True, insertSlidesOnly:=True)
    If newsectioncount = 0 Then
        MsgBox "No top-level bullet points found in the current slide", vbInformation + vbOKOnly, MACROTITLE
    End If
    
    Exit Sub
CreateSlidesFromBulletsErr:
    MsgBox "Please switch to normal or slide view in any presentation." & vbCrLf & _
        "Select a slide which has a bulleted list and try again.", vbExclamation, MACROTITLE
End Sub

Private Function ProcessSlideIntoSections(ByVal sld As Slide, processLowerLevels As Boolean, insertSlidesOnly As Boolean) As Long
    Dim shp As Shape
    Dim sections As Collection
    Dim sectionCount As Long
    
    sectionCount = 0
    For Each shp In sld.Shapes
        If shp.HasTextFrame Then
            Dim tf As TextFrame
            Set tf = shp.TextFrame
            If tf.HasText Then
                Set sections = ProcessTextRange(tf.TextRange, processLowerLevels, insertSlidesOnly)
                sectionCount = sections.Count
                If sectionCount > 0 Then
                    Exit For
                End If
            End If
        End If
    Next
    
    If sectionCount = 0 Then
        ProcessSlideIntoSections = 0
        Exit Function
    End If
    
    Dim currentSection As SectionFromBullets
    For Each currentSection In sections
        currentSection.InsertIntoPresentation sld.Parent
    Next
    
    ProcessSlideIntoSections = sectionCount
End Function

Private Function ProcessTextRange(ByVal tr As TextRange, processLowerLevels As Boolean, insertSlidesOnly As Boolean) As Collection
    Dim i As Long
    Dim para As TextRange
    
    Dim currentIndent As Long
    Dim currentSection As SectionFromBullets
    
    Dim currentSectionTitle As String
    Dim sections As New Collection
    Dim sectionContents As New Collection
    
    For i = 1 To tr.Paragraphs.Count
        Set para = tr.Paragraphs(i)
        If para.ParagraphFormat.Bullet.Visible Then
            If para.Words.Count > 0 And para.IndentLevel = 1 Then
                currentIndent = 1
                currentSectionTitle = para.TrimText
                Set currentSection = NewSection(currentSectionTitle, insertSlidesOnly)
                sections.Add currentSection, currentSectionTitle
            ElseIf para.IndentLevel = 0 Then
                currentIndent = 0
                currentSectionTitle = ""
                Set currentSection = Nothing
            Else
                If para.IndentLevel > currentIndent And Not currentSection Is Nothing Then
                    If processLowerLevels Then
                        currentSection.Contents.Add para
                    End If
                Else
                    Debug.Print MACROTITLE & "[DEBUG]: Bullets seem improperly indented."
                End If
            End If
        End If
    Next
    
    Set ProcessTextRange = sections
End Function

Private Function NewSection(ByVal sectionTitle As String, insertSlidesOnly As Boolean) As SectionFromBullets
    Dim result As SectionFromBullets
    Set result = New SectionFromBullets
    result.Title = sectionTitle
    result.PagesOnly = insertSlidesOnly
    Set NewSection = result
End Function

