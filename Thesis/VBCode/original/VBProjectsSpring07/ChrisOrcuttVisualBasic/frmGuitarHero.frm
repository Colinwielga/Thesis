VERSION 5.00
Begin VB.Form frmGuitarHero 
   BackColor       =   &H00000000&
   Caption         =   "Guitar Hero: Next"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3015
      Left            =   3960
      ScaleHeight     =   2955
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   2160
      Width           =   5535
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
      Height          =   615
      Left            =   7440
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   5190
      Left            =   0
      Picture         =   "frmGuitarHero.frx":0000
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmGuitarHero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmGuitarHero
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmGuitarHero.Hide      'Hides GuitarHero form
    frmIndustryNews.Show    'Shows IndustryNews form
End Sub
'This command opens the GuitarHero.txt file and displays information
'about the game system featured in the picture box
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\GuitarHero.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, GuitarHero(Ctr)
            picResults.Print ; GuitarHero(Ctr)
            Loop
        Close #1
End Sub
