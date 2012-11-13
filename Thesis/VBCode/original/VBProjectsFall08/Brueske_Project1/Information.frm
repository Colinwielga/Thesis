VERSION 5.00
Begin VB.Form Information 
   BackColor       =   &H8000000D&
   Caption         =   "Information"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhome 
      Caption         =   "Home"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdSocial 
      Caption         =   "Social Distance"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H0080FFFF&
      Height          =   4095
      Left            =   1320
      ScaleHeight     =   4035
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdAlienation 
      Caption         =   "Alienation"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Alienation and Social Distance Project
'Results Form
'Kevin Brueske
'Created Nov 3, 2008
'Objective
    'Give some explanations to the concepts being measured
Dim concept(1 To 5) As String

Private Sub cmdalienation_Click()
'Output alienation information
    picoutput.Cls
    picoutput.Print concept(1)
End Sub

Private Sub cmdhome_Click()
'Switch to the home form
    Information.Hide
    Home.Show
End Sub

Private Sub cmdSocial_Click()
'Out social distance information
    picoutput.Cls
    picoutput.Print concept(2)
End Sub

Private Sub Command1_Click()
    Dim ctr As Single
        'File Input, load data into arrays
    Open App.Path & "\info.txt" For Input As #1
        Do Until EOF(1)
            ctr = ctr + 1
            Input #1, concept(ctr)
        Loop
        Close #1
End Sub


