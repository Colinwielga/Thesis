VERSION 5.00
Begin VB.Form frmSlideShow2 
   BackColor       =   &H80000012&
   Caption         =   "Preview Slide 2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNextSlide 
      Caption         =   "Next Slide"
      Height          =   855
      Left            =   7800
      TabIndex        =   2
      Top             =   8160
      Width           =   2775
   End
   Begin VB.CommandButton cmdLastSlide 
      Caption         =   "Previous Slide"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   8160
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   360
      Picture         =   "frmSlideShow2.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   240
      Width           =   9975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   $"frmSlideShow2.frx":B88D
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   360
      TabIndex        =   4
      Top             =   6480
      Width           =   10575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   $"frmSlideShow2.frx":B930
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   5040
      Width           =   10215
   End
End
Attribute VB_Name = "frmSlideShow2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Solar Project (Final Project 1.VBP)
'frmSlideShow2 (frmDG.frm)
'Dan Gregus
'3/27/08
'Objective: Create a brief slideshow describing the Solar Project

Private Sub cmdLastSlide_Click()
    frmSlideShow2.Visible = False
    frmSlideShow1.Visible = True
End Sub

Private Sub cmdNextSlide_Click()
    frmSlideShow2.Visible = False
    frmSlideShow3.Visible = True
End Sub

