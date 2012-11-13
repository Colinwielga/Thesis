VERSION 5.00
Begin VB.Form Trivia 
   Caption         =   "Trivia"
   ClientHeight    =   12285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14460
   BeginProperty Font 
      Name            =   "Footlight MT Light"
      Size            =   8.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "Trivia.frx":0000
   ScaleHeight     =   12285
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Information"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   20
      Top             =   11520
      Width           =   4575
   End
   Begin VB.CommandButton cmdL4 
      Caption         =   "Spanish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      TabIndex        =   19
      Top             =   10320
      Width           =   1815
   End
   Begin VB.CommandButton cmdL2 
      Caption         =   "German"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   18
      Top             =   10320
      Width           =   1815
   End
   Begin VB.CommandButton cmdL3 
      Caption         =   "British"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   17
      Top             =   10320
      Width           =   1815
   End
   Begin VB.CommandButton cmdT4 
      Caption         =   "Mancomunado Trophy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      TabIndex        =   16
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdT2 
      Caption         =   "Copa del Rey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   15
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdT3 
      Caption         =   "La Liga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   14
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdL1 
      Caption         =   "American(MLS)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   13
      Top             =   10320
      Width           =   1815
   End
   Begin VB.CommandButton cmdK2 
      Caption         =   "Cristiano Obasi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdK3 
      Caption         =   "Mesut Ozil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   11
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdK4 
      Caption         =   "Raul Albiol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdT1 
      Caption         =   "UEFA Cup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   9
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdK1 
      Caption         =   "Iker Casillas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdRooney 
      Caption         =   "Wayne Rooney"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdAlonso 
      Caption         =   "Xabi Alonso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCarvalho 
      Caption         =   "Ricardo Carvalho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdronaldo 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cristiano Ronaldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      MaskColor       =   &H008080FF&
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblq2 
      BackColor       =   &H0000FFFF&
      Caption         =   "2) What is the name of Real Madrid's Goalkeeper?"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   24
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   14535
   End
   Begin VB.Label lblq3 
      BackColor       =   &H00FF0000&
      Caption         =   "3) What trophy did Real Madrid win the most?"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   24
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   5880
      Width           =   14535
   End
   Begin VB.Label lblq4 
      BackColor       =   &H000080FF&
      Caption         =   "4) What league is Real Madrid located in?"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   24
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   9000
      Width           =   14535
   End
   Begin VB.Label lblq1 
      BackColor       =   &H000000FF&
      Caption         =   "1) Which of these players is not a member of Real Madrid's Roster?"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   26.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "Trivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form allows the user to show how much he or she has learned from the previous information
'or simply test their knowledge of what they know regarding Real Madrid
Option Explicit
Private Sub cmdAlonso_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdBack_Click()
Information.Show
Statistics.Hide
PlayersStat.Hide
OpenPage.Hide
Me.Hide
End Sub

Private Sub cmdCarvalho_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdK1_Click()
MsgBox ("Correct")
End Sub

Private Sub cmdK2_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdK3_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdK4_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdL1_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdL2_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdL3_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdL4_Click()
MsgBox ("Correct")
End Sub

Private Sub cmdronaldo_Click()
MsgBox ("Wrong")

End Sub

Private Sub cmdRooney_Click()
MsgBox ("Correct")
End Sub

Private Sub cmdT1_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdT2_Click()
MsgBox ("Wrong")
End Sub

Private Sub cmdT3_Click()
MsgBox ("Correct")
End Sub

Private Sub cmdT4_Click()
MsgBox ("Wrong")
End Sub
