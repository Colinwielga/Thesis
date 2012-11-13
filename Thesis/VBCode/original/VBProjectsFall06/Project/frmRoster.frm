VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H00000080&
   Caption         =   "History of Gopher Hockey"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   615
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   7440
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   4215
      Left            =   3840
      Picture         =   "frmRoster.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   3000
      Width           =   5415
   End
   Begin VB.CommandButton cmdHobey 
      Caption         =   "Hobey Baker Award Winners"
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   8175
      Left            =   120
      Picture         =   "frmRoster.frx":F9F8
      ScaleHeight     =   8115
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdChamps 
      Caption         =   "NCAA Championships"
      Height          =   615
      Index           =   1
      Left            =   4080
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdAllAmerican 
      Caption         =   "Past All-Americans"
      Height          =   615
      Index           =   0
      Left            =   6840
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label lblHistory 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "History of Gopher Hockey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   5475
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmHistory
'Cole and John
'10/30/06
'Objective: The objective of this form is to allow the user to choose from viewing
'(1) Hobey Baker award winners, (2) NCAA championships, and (3) All-Americans.  The
'user can access this information by clicking on the appropriate command button.

Option Explicit
Private Sub cmdAllAmerican_Click(Index As Integer)
    frmAllAmerican.Visible = True       'makes visible the AllAmerican form, hides the History form
    frmHistory.Visible = False
End Sub

Private Sub cmdChamps_Click(Index As Integer)
    frmChamps.Visible = True            'same as above
    frmHistory.Visible = False
End Sub

Private Sub cmdHobey_Click()
    frmHobeyBaker.Visible = True
    frmHistory.Visible = False
End Sub

Private Sub cmdHome_Click(Index As Integer)
    frmMain.Visible = True          'allows user to go back
    frmHistory.Visible = False
End Sub
