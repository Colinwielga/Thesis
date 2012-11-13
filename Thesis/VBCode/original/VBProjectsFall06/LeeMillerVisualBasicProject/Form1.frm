VERSION 5.00
Begin VB.Form Top10 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quitcmd 
      BackColor       =   &H000000C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   615
   End
   Begin VB.CommandButton WRcmd 
      BackColor       =   &H00008080&
      Caption         =   "Wide Recievers"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   840
      Picture         =   "Form1.frx":1A7292
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton RBcmd 
      BackColor       =   &H00800000&
      Caption         =   "Running Backs"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   840
      Picture         =   "Form1.frx":1B832C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "Quarterbacks"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   840
      Picture         =   "Form1.frx":1C5A96
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Click on a player"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Top 10 in 2005"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Top10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name = Top 10 in 2005'
'Form Name = Top 10 in 2005'
'Lee Miller'
'October 31, 2006'
'The purpose of this form is for the user to reach 3 other forms that will allow you
'to sort and search through the top 10 Quarterbacks, Running Backs, and Wide Recievers
'from 2005'

Private Sub Command1_Click()
'This command button takes you to the Quarterback form'
QB.Show
Top10.Hide
QB.Visible = True
End Sub

Private Sub Quitcmd_Click()
End
End Sub

Private Sub RBcmd_Click()
'This command button takes you to the Running Backs form'
RB.Show
Top10.Hide
RB.Visible = True
End Sub

Private Sub WRcmd_Click()
'This command button takes you to the Wide Recievers form'
WR.Show
Top10.Hide
WR.Visible = True
End Sub
