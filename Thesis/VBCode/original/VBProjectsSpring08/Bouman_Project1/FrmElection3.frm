VERSION 5.00
Begin VB.Form FrmElection3 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton CmdViewPopVote 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Results By Popular Vote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton CmdViewDelegates 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Results by Delegates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   2040
      Picture         =   "FrmElection3.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   2040
      Width           =   5535
   End
End
Attribute VB_Name = "FrmElection3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Election Project
'FrmElection3
'Ian Bouman
'Written on 3/14
'This objective of this form is to direct the user to the results
'of his choice - delegate results or popular vote results.
Private Sub CmdViewDelegates_Click()
FrmElection3.Hide
FrmDelegates.Show
End Sub

Private Sub CmdViewPopVote_Click()
FrmElection3.Hide
FrmPopVote.Show
End Sub

Private Sub Command3_Click()
FrmElection3.Hide
FrmElection1.Show
End Sub

Private Sub Command4_Click()
End
End Sub
