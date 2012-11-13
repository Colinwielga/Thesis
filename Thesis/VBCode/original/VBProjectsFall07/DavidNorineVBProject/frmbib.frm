VERSION 5.00
Begin VB.Form frmbib 
   BackColor       =   &H0000FF00&
   Caption         =   "Bibliography"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   7560
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdbib1 
      Caption         =   "Load Bibliography"
      Height          =   975
      Left            =   7440
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.PictureBox picresults 
      Height          =   9135
      Left            =   360
      ScaleHeight     =   9075
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   4230
      Left            =   7320
      Picture         =   "frmbib.frx":0000
      Top             =   4320
      Width           =   5700
   End
End
Attribute VB_Name = "frmbib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdbib1_Click()
Dim bib(1 To 15) As String
Dim CTR As Integer, pos As Integer




Open App.Path & "\bibliography.txt" For Input As #1

CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, bib(CTR) 'puts the list of sites into an array
    Loop

    Close #1 'closes the bibliography.txt file

    For pos = 1 To CTR
        picresults.Print bib(pos) 'displays the list of sites
        picresults.Print "                                                                                               "
    Next pos
End Sub

Private Sub cmdreturn_Click()
'this subroutine goes back to the main menu
frmbib.Hide
frmmain.Show

End Sub
