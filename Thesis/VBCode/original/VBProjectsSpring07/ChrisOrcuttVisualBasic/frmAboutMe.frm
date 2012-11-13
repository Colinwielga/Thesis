VERSION 5.00
Begin VB.Form frmAboutMe 
   Caption         =   "About Me"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4080
      ScaleHeight     =   3315
      ScaleWidth      =   4470
      TabIndex        =   3
      Top             =   960
      Width           =   4530
   End
   Begin VB.CommandButton cmdMyGames 
      Caption         =   "My Games"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblChris 
      BackStyle       =   0  'Transparent
      Caption         =   "This is me:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   3735
   End
   Begin VB.Image ImageAboutMe 
      Height          =   6015
      Left            =   0
      Picture         =   "frmAboutMe.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmAboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form tells a little about myself and where I go to school
'Clicking on the image in the window will bring up an "About Me" section
Option Explicit
Private Sub cmdMyGames_Click()
    frmAboutMe.Hide     'Hides AboutMe form
    frmMyGames.Show     'Shows MyGames form
End Sub
Private Sub cmdReturn_Click()
    frmAboutMe.Hide     'Hides AboutMe form
    frmSelectWant.Show  'Shows SelectWant form
End Sub
'This Displays my "About Me" in the picture box
Private Sub ImageAboutMe_Click()
    Dim Ctr As Integer
        Open App.Path & "\AboutMe.txt" For Input As #1      'Opens txt document for display
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, AboutMe(Ctr)
            picResults.Print ; AboutMe(Ctr)
            Loop
        Close #1
End Sub
