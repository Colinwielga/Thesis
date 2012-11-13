VERSION 5.00
Begin VB.Form frmHardCaseSolve 
   BackColor       =   &H00000000&
   Caption         =   "Hard Case Solution"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdreturn2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to case files page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to title Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2415
   End
   Begin VB.PictureBox picone 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   6480
      Picture         =   "frmHardCaseSolve.frx":0000
      ScaleHeight     =   2295
      ScaleWidth      =   3135
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Congratulations you have concluded that the cause of death was autoerotic manslaughter."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4200
      TabIndex        =   0
      Top             =   480
      Width           =   7935
   End
End
Attribute VB_Name = "frmHardCaseSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This simply is a way to go form. It has no calculations
'however i did leave a picture box on the form that always appears when the form is
'activated. So that is the case closed picture you see. You can't do much like i said
'here all you can do is go back to previous forms.


Private Sub cmdReturn_Click()
'Takes the User back to the title screen so they can begin agian with a new
'name or quit the program
    frmHardCaseSolve.Hide
    frmTitleScreen.Show
    
End Sub

Private Sub cmdreturn2_Click()
'takes the user back to the case files form so they can do the other cases
'available for them.
    frmHardCaseSolve.Hide
    frmCasefiles.Show
End Sub
