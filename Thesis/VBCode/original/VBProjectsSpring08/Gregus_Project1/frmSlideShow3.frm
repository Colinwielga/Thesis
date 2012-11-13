VERSION 5.00
Begin VB.Form frmSlideShow3 
   BackColor       =   &H80000012&
   Caption         =   "Preview Slide 3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMainPage 
      Caption         =   "Back to Main Page"
      Height          =   1215
      Left            =   7320
      TabIndex        =   2
      Top             =   7560
      Width           =   3375
   End
   Begin VB.CommandButton cmdPreviousSlide 
      Caption         =   "Previous Slide"
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   7560
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   360
      Picture         =   "frmSlideShow3.frx":0000
      ScaleHeight     =   4755
      ScaleWidth      =   10635
      TabIndex        =   0
      Top             =   240
      Width           =   10695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   $"frmSlideShow3.frx":AEC5
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   840
      TabIndex        =   3
      Top             =   5520
      Width           =   9855
   End
End
Attribute VB_Name = "frmSlideShow3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Solar Project (Final Project 1.VBP)
'frmSlideShow3 (frmDG.frm)
'Dan Gregus
'3/27/08
'Objective: Create a brief slideshow describing the Solar Project


Private Sub CmdMainPage_Click()
    frmSlideShow3.Visible = False
    frmSolarProject.Visible = True
    
End Sub

Private Sub cmdPreviousSlide_Click()
    frmSlideShow3.Visible = False
    frmSlideShow2.Visible = True
End Sub

