VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H80000007&
   Caption         =   "Menu"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8790
   ScaleWidth      =   13260
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   7320
      Width           =   3975
   End
   Begin VB.CommandButton cmdDebt 
      Caption         =   "Which Clubs Are In Debt? "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9480
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdRugby 
      Caption         =   "Women's Rugby Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdFusion 
      Caption         =   "Cultural Fusion Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdDems 
      Caption         =   "College Democrats Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdSpanishClub 
      Caption         =   "Spanish Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdComm 
      Caption         =   "Communication Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblName 
      Caption         =   "Created by Jody Roers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblAuditor 
      Caption         =   "  Club Auditor Aid"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblClub 
      BackColor       =   &H0080FFFF&
      Caption         =   "Please Choose A Club:"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblSenate 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CSB SENATE         2003-04        CLUB AUDITOR"
      BeginProperty Font 
         Name            =   "SchoolBoy"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Club Auditing Aid (VBProject.vbp)
'Form Name: Menu (VBMenu.frm)
'Author: Jody Roers
'Date Written: 27 October 2003
'Purpose: To serve as a map to allow the user to choose a path.  User directs which form user
'would like to utalize.
Public Path As String

Private Sub cmdComm_Click()
Menu.Hide 'go to Comm form
Comm.Show
End Sub

Private Sub cmdDebt_Click()
Menu.Hide 'go to Debt form
Debt.Show
End Sub

Private Sub cmdDems_Click()
Menu.Hide 'go to College Dems form
Dems.Show
End Sub

Private Sub cmdFusion_Click()
Menu.Hide 'got to Fusion form
Fusion.Show
End Sub

Private Sub cmdQuit_Click()
Dim Name As String
Name = InputBox("Please Enter Your Name", "Enter")
Name = Name & " Rocks!!!"
MsgBox Name, , "Thanks for Coming"
End 'quit program
End Sub

Private Sub cmdRugby_Click()
Menu.Hide 'go to rugby form
Rugby.Show
End Sub

Private Sub cmdSpanishClub_Click()
Menu.Hide 'to to spanish form
Spanish.Show
End Sub


Private Sub Form_Load()
Path = "N:\CS130\handin\Jody_Roers\"
End Sub
