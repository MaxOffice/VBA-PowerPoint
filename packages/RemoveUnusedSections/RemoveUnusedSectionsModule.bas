Attribute VB_Name = "RemoveUnusedSectionsModule"
Option Explicit

Private Const MACROTITLE As String = "Remove Unused Sections"

Public Sub RemoveUnusedSections()

    ' 10 Feb 2025
    ' Nitin Paranjape
    ' Deletes empty sections
    Dim ap As Presentation
    Dim delCnt As Integer, x As Integer
    
    On Error GoTo RemoveUnusedSectionsError
    
    Set ap = ActivePresentation
    If ap Is Nothing Then
        MsgBox "Please open a presentation, and try again.", _
                vbOKOnly, MACROTITLE
        Exit Sub
    End If

    With ap.SectionProperties
        
        If .Count = 0 Then
            MsgBox "There are no sections in this presentation.", _
                    vbInformation + vbOKOnly, MACROTITLE
            Exit Sub
        End If
    
        Dim i As Integer
        x = .Count
        For i = x To 1 Step -1
            
            If .SlidesCount(i) = 0 Then
                .Delete i, False
                
                delCnt = delCnt + 1
            End If
        Next i

        If delCnt <= 0 Then
            MsgBox "No empty sections found.", _
                    vbInformation + vbOKOnly, MACROTITLE
        Else
            MsgBox "Deleted " & delCnt & " sections.", _
                    vbInformation + vbOKOnly, MACROTITLE
        End If

   End With
   
   Exit Sub
RemoveUnusedSectionsError:
    MsgBox "Error while performing operation: " & Err.Description & vbCrLf & _
           "Please try again.", vbOKOnly + vbExclamation, MACROTITLE
End Sub

