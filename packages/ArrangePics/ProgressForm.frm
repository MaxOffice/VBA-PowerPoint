VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Working ..."
   ClientHeight    =   870
   ClientLeft      =   2120
   ClientTop       =   2460
   ClientWidth     =   4570
   OleObjectBlob   =   "ProgressForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private mPicCount As Integer
Private mCurrentPic As Integer

Public Property Get PicCount() As Integer
    PicCount = mPicCount
End Property

Public Property Let PicCount(value As Integer)
    mPicCount = value
    
    sbProgress.max = mPicCount
End Property

Public Property Get CurrentPic() As Integer
    CurrentPic = mCurrentPic
End Property

Public Property Let CurrentPic(value As Integer)
    mCurrentPic = value
    
    sbProgress.value = mCurrentPic
    Caption = "Processing picture " & mCurrentPic & "/" & mPicCount
End Property

