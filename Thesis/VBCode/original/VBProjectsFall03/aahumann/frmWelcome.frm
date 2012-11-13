VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbxNfl 
      Height          =   3375
      Left            =   4320
      Picture         =   "frmWelcome.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.PictureBox pbxLt 
      Height          =   3975
      Left            =   7560
      Picture         =   "frmWelcome.frx":3956
      ScaleHeight     =   3915
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   960
      Width           =   2895
   End
   Begin VB.PictureBox pbxMoss 
      Height          =   3855
      Left            =   360
      Picture         =   "frmWelcome.frx":7255
      ScaleHeight     =   3795
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7320
      TabIndex        =   2
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Click Here to Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      TabIndex        =   1
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to NFL Fantasy Football 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : NFL (nfl.vbp)
'Form Name : frmWelcome (Welcome.frm)
'Project Name: NFL (nfl.vbp)
'Form Name: frmWelcome (frmWelcome.frm)

'Author: Andy Humann
'Date Written: Oct. 28th, 2003
'Purpose of Project: to allow user to view my current fantasy football info
                    'and allows user to calculate fantasy points for the week
'Purpose of Form: To welcome the user to the program
            'allows user to enter or quit the program

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit





Private Sub cmdEnter_Click()
'used to switch to the next form when cmdEnter is clicked

    frmPick.Visible = True
    frmWelcome.Visible = False
End Sub

Private Sub cmdQuit_Click()
'quits program
    End
End Sub

