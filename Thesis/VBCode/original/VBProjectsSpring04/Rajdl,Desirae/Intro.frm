VERSION 5.00
Begin VB.Form Introform 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbegin 
      Caption         =   "Begin"
      Height          =   975
      Left            =   6720
      TabIndex        =   1
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label lblname2 
      BackColor       =   &H00FF8080&
      Caption         =   "By: Desirae Rajdl"
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"Intro.frx":0000
      Height          =   3135
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "Introform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbegin_Click()
'JEC Voting Analysis  (JECvoting)
'Introductary Help Page (Introform)
'By: Desirae Rajdl
'Written: March 10, 2004
'This form was written just to explain the basics of what the other
'form's function is and a little about how to work it.
Introform.Hide
Analyzeform.Show
End Sub
