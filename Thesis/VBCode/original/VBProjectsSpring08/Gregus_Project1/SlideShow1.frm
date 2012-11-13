VERSION 5.00
Begin VB.Form frmSlideShow1 
   BackColor       =   &H80000007&
   Caption         =   "Preview Slide 1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNextSlide 
      Caption         =   "Next Slide"
      Height          =   975
      Left            =   7680
      TabIndex        =   2
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdMainPage 
      BackColor       =   &H8000000E&
      Caption         =   "Go back to the main page"
      Height          =   975
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   240
      Picture         =   "SlideShow1.frx":0000
      Top             =   360
      Width           =   12000
   End
   Begin VB.Label ltl 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   $"SlideShow1.frx":170DE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2775
      Left            =   600
      TabIndex        =   0
      Top             =   5280
      Width           =   8895
   End
End
Attribute VB_Name = "frmSlideShow1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Solar Project (Final Project 1.VBP)
'frmSlideShow1 (frmDG.frm)
'Dan Gregus
'3/27/08
'Objective: Create a brief slideshow describing the Solar Project
Private Sub CmdMainPage_Click()
    frmSlideShow1.Visible = False
    frmSolarProject.Visible = True
End Sub

Private Sub cmdNextSlide_Click()
    frmSlideShow1.Visible = False
    frmSlideShow2.Visible = True
End Sub
