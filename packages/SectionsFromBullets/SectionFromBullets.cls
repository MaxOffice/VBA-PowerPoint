VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SectionFromBullets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sectionTitle As String
Private m_sectionContents As Collection

Public Property Get Title() As String
    Title = m_sectionTitle
End Property

Friend Property Let Title(ByVal value As String)
    m_sectionTitle = value
End Property

Public Property Get Contents() As Collection
    Set Contents = m_sectionContents
End Property

Public Property Get ContentCount() As Long
    ContentCount = m_sectionContents.Count
End Property

Private Sub Class_Initialize()
    Set m_sectionContents = New Collection
End Sub

Public Sub InsertIntoPresentation(ByVal pres As Presentation)
    Dim lastSlideIndex As Long
    With pres
        lastSlideIndex = .Slides.Count
        
        ' Add a section header slide
        Dim titleSlide As Slide
        Set titleSlide = .Slides.AddSlide(lastSlideIndex + 1, .SlideMaster.CustomLayouts(ppLayoutTitle))
        titleSlide.Shapes.Title.TextFrame.TextRange.Text = m_sectionTitle
        
        ' Add section break just before it
        .SectionProperties.AddBeforeSlide lastSlideIndex + 1, m_sectionTitle
        
        If ContentCount > 0 Then
            Dim contentSlide As Slide
            Set contentSlide = .Slides.AddSlide(lastSlideIndex + 2, .SlideMaster.CustomLayouts(ppLayoutText))
            
            Dim tr As TextRange
            Set tr = contentSlide.Shapes.Placeholders(ppPlaceholderBody).TextFrame.TextRange
            
            Dim para As TextRange
            For Each para In m_sectionContents
                tr.Paragraphs.InsertAfter(para).IndentLevel = para.IndentLevel - 1
            Next
        End If
    End With
End Sub
