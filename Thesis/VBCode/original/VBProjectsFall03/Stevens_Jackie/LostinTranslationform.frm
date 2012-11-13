VERSION 5.00
Begin VB.Form LostinTranslation 
   BackColor       =   &H00800080&
   Caption         =   "Lost in Translation"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdBackInfo 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Movie Info"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdBackMain 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   240
      Picture         =   "LostinTranslationform.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
      Caption         =   "Courtesy of Hollywood.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   6720
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   $"LostinTranslationform.frx":A4AC
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2775
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      Width           =   5535
   End
End
Attribute VB_Name = "LostinTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MovieProject (MoveProject.vbp)
'Form Name: LostinTranslation (LostinTranslationform.frm)
'Author: Jackie Stevens
'Date Written: 10/20/03
'Purpose: 1. To display information about the particular movie in order
            'help the user make a movie decision
         '2. To provide links to move to other forms in the program.

Option Explicit

Private Sub cmdBackInfo_Click()
    'Go back to movie info form
MovieInfo.Show
LostinTranslation.Hide
End Sub

Private Sub cmdBackMain_Click()
'Go back to main movie form
MovieMain.Show
LostinTranslation.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub
