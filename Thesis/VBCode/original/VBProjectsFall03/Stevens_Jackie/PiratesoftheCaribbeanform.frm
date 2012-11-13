VERSION 5.00
Begin VB.Form PiratesoftheCaribbean 
   BackColor       =   &H00800080&
   Caption         =   "Pirates of the Caribbean"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11385
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   240
      Picture         =   "PiratesoftheCaribbeanform.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
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
      Left            =   600
      TabIndex        =   2
      Top             =   6720
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   $"PiratesoftheCaribbeanform.frx":B0FE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5535
      Left            =   4920
      TabIndex        =   1
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "PiratesoftheCaribbean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MovieProject (MoveProject.vbp)
'Form Name: PiratesoftheCaribbean (PiratesoftheCaribbeanform.frm)
'Author: Jackie Stevens
'Date Written: 10/20/03
'Purpose: 1. To display information about the particular movie in order
            'help the user make a movie decision
         '2. To provide links to move to other forms in the program.

Option Explicit

Private Sub cmdBackInfo_Click()
    'Go back to movie info form
MovieInfo.Show
PiratesoftheCaribbean.Hide
End Sub

Private Sub cmdBackMain_Click()
'Go back to main movie form
MovieMain.Show
PiratesoftheCaribbean.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub