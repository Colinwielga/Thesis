VERSION 5.00
Begin VB.Form ColdCreekManor 
   BackColor       =   &H00800080&
   Caption         =   "Cold Creek Manor"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11985
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
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
      Top             =   6240
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   120
      Picture         =   "ColdCreekManorform.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   240
      Width           =   4335
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   $"ColdCreekManorform.frx":9F68
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
      Height          =   4575
      Left            =   5160
      TabIndex        =   1
      Top             =   960
      Width           =   5175
   End
End
Attribute VB_Name = "ColdCreekManor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MovieProject (MoveProject.vbp)
'Form Name: ColdCreekManor (ColdCreekManorform.frm)
'Author: Jackie Stevens
'Date Written: 10/20/03
'Purpose: 1. To display information about the particular movie in order
            'help the user make a movie decision
         '2. To provide links to move to other forms in the program.
Option Explicit
Private Sub cmdBackInfo_Click()
    'Go to info page
MovieInfo.Show
ColdCreekManor.Hide
End Sub
Private Sub cmdBackMain_Click()
    'Go to Main page
MovieMain.Show
ColdCreekManor.Hide
End Sub
Private Sub cmdQuit_Click()
    'Quits program
End
End Sub
