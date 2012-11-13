VERSION 5.00
Begin VB.Form frmRiver 
   BackColor       =   &H00FF0000&
   Caption         =   "River!"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Forward Ho!"
      Height          =   735
      Left            =   4320
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nice Work! You killed them all! On to the people that need help!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image img3 
      Height          =   2370
      Left            =   1680
      Picture         =   "frmRiver.frx":0000
      Top             =   1200
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Image img2 
      Height          =   2370
      Left            =   6720
      Picture         =   "frmRiver.frx":1C9F
      Top             =   1440
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Oh no! There's sharks in the River!  Quick Click on each one and cast a spell on it!!! "
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
   Begin VB.Image img1 
      Height          =   2370
      Left            =   480
      Picture         =   "frmRiver.frx":393E
      Top             =   2880
      Width           =   2610
   End
End
Attribute VB_Name = "frmRiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'In this form, the user clicks on the sharks to kill them and
'make them disappear.

Private Sub CmdNext_Click()
 img1.Visible = True
 img2.Visible = False
 img3.Visible = False
 lbl2.Visible = False
 cmdNext.Visible = False
 frmRiver.Hide
 frmPeople.Show
 
End Sub

Private Sub cmdQuit_Click()
    End
    
End Sub

Private Sub img1_Click()
 img1.Visible = False
 img2.Visible = True
 
End Sub

Private Sub img2_Click()
 img2.Visible = False
 img3.Visible = True
 
End Sub

Private Sub img3_Click()
 img3.Visible = False
 lbl2.Visible = True
 cmdNext.Visible = True
 
End Sub
