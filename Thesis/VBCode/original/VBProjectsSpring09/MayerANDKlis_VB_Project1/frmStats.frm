VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H000000C0&
   Caption         =   "Team Statistics"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "@GungsuhChe"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
      Width           =   3015
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to main page"
      Height          =   735
      Left            =   5160
      TabIndex        =   6
      Top             =   9240
      Width           =   2895
   End
   Begin VB.CommandButton CmdAnswer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Am I correct?"
      BeginProperty Font 
         Name            =   "@GungsuhChe"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   3135
   End
   Begin VB.TextBox txtanswer 
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   7080
      Width           =   3975
   End
   Begin VB.CommandButton CmdFielders 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minnesota Twins Fielders Statistics"
      BeginProperty Font 
         Name            =   "@GungsuhChe"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdPitchers 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minnesota Twins Pitchers Statistics"
      BeginProperty Font 
         Name            =   "@GungsuhChe"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   2040
      Picture         =   "frmStats.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What is the name of the Minnesota Twins manager pictured above?"
      BeginProperty Font 
         Name            =   "@GungsuhChe"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   7200
      Width           =   4695
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Minnesota Twins
'FrmStats
'Sarah Mayer and Jake Klis
'Written on 03/22/09
'The goal of this form is to ask you a basic trivia question and also act as a guide to other
'forms where you can find statistical information about the Minnesota Twins

'This button sees if the the input from the user matches the correct answer. If it is correct
'the Trivia counter is incremented by one.
Private Sub CmdAnswer_Click()
Dim TextBox As String
TextBox = TxtAnswer
If TextBox = "Ron Gardenhire" Then
    MsgBox ("You are so correct!")
    TriviaCtr = TriviaCtr + 1
    Else
    MsgBox ("Have you ever even seen a twins game? The correct answer is Ron Gardenhire!")
    End If
    
    
End Sub

Private Sub Cmdback_Click()
frmStats.Hide
FrmMain.Show
End Sub

Private Sub CmdFielders_Click()
FrmFielders.Show
frmStats.Hide
End Sub

Private Sub CmdPitchers_Click()
frmPitchers.Show
frmStats.Hide
End Sub

Private Sub Cmdreturn_Click()
frmStats.Hide
FrmMain.Show

End Sub
