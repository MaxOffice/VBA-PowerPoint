Attribute VB_Name = "pptControlModule"
Option Explicit
Option Base 1

Public Sub PresentCurrentSlide()

    ' Created by Dr Nitin Paranjape on 13 Aug 2021
    '
    ' This assumes that you have dual monitor setup OR
    ' if you are using a projector in extended monitor mode to display the presentations
        
    ' This macro is run when you select a slide and you want that slide to be shown
    ' This is useful when you have multiple PPTs running and the presenter view is ON
    ' Why is this useful?
        ' Scenario
        ' Three ppts running
        ' Presenter view shows only ONE ppt at a time
        ' Let us say I am showing the 5th slide in all the presentations
        ' Currently slide 5 of ppt1 is running
        ' Now I want to show slide 7 of ppt2
        ' In order to do this, I will have to
            ' Hover on the PowerPoint icon in taskbar
            ' Locate the ppt2 presentation thumbnail and click on it
            ' At this stage, it is showing slide 5 of ppt2
            ' Then I have to go to slide view in Presenter View and move to slide 7 of ppt2
        ' The problem is that, I wanted to show the 7th slide of ppt2 WITHOUT showing its currently projected slide
        ' This is impossible to do
        ' That is why this macro is useful
        
    ' With the macro running, I can keep any number of ppts
    ' Click on any slide in edit mode
    ' Run the macro, it will directly show the desired slide
    
    
    Dim sldx As Integer
    Dim ap As Presentation
    Dim wn As SlideShowWindow
    Dim apn As String
    Dim isrunning As Boolean
    
    
    'Running slide show flag
    isrunning = False
    
    Set ap = ActivePresentation
    
    apn = ap.Name
    'Check if the current presentation in edit mode running
    For Each wn In Application.SlideShowWindows
    
       If wn.Presentation.Name = apn Then
           
           isrunning = True
           Set wn = Nothing
       
           Exit For
       
           
       End If
       
    
    Next 'each wn

    'Get the index of current slide open in edit mode in the active presentation
    'This is the slide we want to show next
    
    sldx = ActiveWindow.View.Slide.SlideIndex
    Set ap = ActiveWindow.Presentation
    
    'Check slideshow is running for the active presentation in edit mode
    
    If isrunning Then
       
       'Change the current slide and activate
       ap.SlideShowWindow.View.GotoSlide sldx, msoTrue
       
       'Activate the slide show
       ap.SlideShowWindow.Activate
    
    Else 'isrunning
    
       'Run this slide show from current slide in edit view
       
       CommandBars.ExecuteMso "SlideShowFromCurrent"
       
    End If 'isrunning
    
    
End Sub

Public Sub EndAllShows()
    
    ' Stops all running slide shows
    
    With Application
    
        Dim f As Integer
        
        ' Iterate running presentations
        
        For f = 1 To .SlideShowWindows.Count
        
            .SlideShowWindows(1).Activate
            ' Close slide show
            .SlideShowWindows(1).View.Exit
            
            ' Maximize the base presentation
            .ActiveWindow.WindowState = ppWindowMaximized
            
        Next
    
    End With
    
End Sub



