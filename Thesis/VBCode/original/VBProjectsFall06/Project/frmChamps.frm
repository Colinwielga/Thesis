VERSION 5.00
Begin VB.Form frmChamps 
   BackColor       =   &H00000080&
   Caption         =   "NCAA Champions"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Index           =   1
      Left            =   2640
      TabIndex        =   8
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   735
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   6840
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   3600
      Picture         =   "frmChamps.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   1320
      Width           =   6015
   End
   Begin VB.CommandButton cmd2002 
      Caption         =   "2002 NCAA Champions"
      Height          =   735
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmd1979 
      Caption         =   "1979 NCAA Champions"
      Height          =   735
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmd1976 
      Caption         =   "1976 NCAA Champions"
      Height          =   735
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmd1974 
      Caption         =   "1974 NCAA Champions"
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmd2003 
      Caption         =   "2003 NCAA Champions"
      Height          =   735
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblChamps 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "NCAA Championships"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   4830
   End
End
Attribute VB_Name = "frmChamps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmChamps
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the user with the opportunity
'to view past national championship teams. The user can select the year they wish
'to view by clicking on the appropriate command button.


Option Explicit


Private Sub cmd1974_Click(Index As Integer)
    frm1974.Visible = True      'makes visible the 1974 Champions form, and hides the current form
    frmChamps.Visible = False
End Sub

Private Sub cmd1976_Click(Index As Integer)
    frm1976.Visible = True      'same as above
    frmChamps.Visible = False
End Sub

Private Sub cmd1979_Click(Index As Integer)
    frm1979.Visible = True
    frmChamps.Visible = False
End Sub

Private Sub cmd2002_Click(Index As Integer)
    frm2002.Visible = True
    frmChamps.Visible = False
End Sub

Private Sub cmd2003_Click(Index As Integer)
    frm2003.Visible = True
    frmChamps.Visible = False
End Sub

Private Sub cmdBack_Click(Index As Integer)
    frmChamps.Visible = False
    frmHistory.Visible = True
End Sub

Private Sub cmdHome_Click(Index As Integer)
    frmChamps.Visible = False
    frmMain.Visible = True
End Sub
