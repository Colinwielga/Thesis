VERSION 5.00
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   Caption         =   "Display of Stats"
   ClientHeight    =   6675
   ClientLeft      =   3135
   ClientTop       =   2760
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000080FF&
      Caption         =   "Back"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      Height          =   4455
      Left            =   960
      ScaleHeight     =   4395
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form  to display to the user three different ways of sorting the information either by date, alphabetically, or rating.
Public Sub showStats()
Dim I As Integer
picResults.Cls  'Clear the picture box
For I = 1 To Counter  'the tab 60 feature allows you to begin printing movierating I at the 60th character place.
    picResults.Print MovieRelease(I), MovieName(I); Tab(60); MovieRating(I)
Next I
End Sub


Private Sub cmdBack_Click()
frmDisplay.Hide     'Allows user to go from Display form to Stats form
frmStats.Show
End Sub

