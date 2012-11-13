VERSION 5.00
Begin VB.Form frmFishTanks 
   BackColor       =   &H00FF0000&
   Caption         =   "Fish Tanks"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdTen 
      Caption         =   "Ten Gallon"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   12
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdFive 
      Caption         =   "Five Gallon"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   11
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdTwo 
      Caption         =   "Two Gallon"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmd5Gall 
      BackColor       =   &H0000C000&
      Caption         =   "Five Gallon  ($50.00)"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1414
      Index           =   2
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmd10Gall 
      BackColor       =   &H0000C000&
      Caption         =   "Ten Gallon  ($75.00)"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1414
      Index           =   0
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmd2Gall 
      BackColor       =   &H0000C000&
      Caption         =   "Two Gallon  ($20.00)"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1414
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton cmdNoThanks 
      Caption         =   "No Thank You!"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      TabIndex        =   2
      Top             =   9000
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11040
      TabIndex        =   1
      Top             =   9000
      Width           =   3975
   End
   Begin VB.PictureBox PicResults 
      Height          =   5055
      Left            =   4320
      ScaleHeight     =   4995
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Label LabInstructions3 
      BackColor       =   &H00FF0000&
      Caption         =   "Click the type of tank you wish to purchase:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Index           =   1
      Left            =   10800
      TabIndex        =   9
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label LabInstructions2 
      BackColor       =   &H00FF0000&
      Caption         =   "**The tanks include filters and lights.  "
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   7200
      Width           =   6375
   End
   Begin VB.Label LabInstructions 
      BackColor       =   &H00FF0000&
      Caption         =   "To view the tanks click on the buttons below:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1095
      Left            =   600
      TabIndex        =   7
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label LabFishTanks 
      BackColor       =   &H00FF0000&
      Caption         =   "Fish Tanks"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   48
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   1080
      TabIndex        =   4
      Top             =   8520
      Width           =   4935
   End
End
Attribute VB_Name = "frmFishTanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmFishTanks
'Author: Scott Sand and Kate Sand
'Date Written: March 10, 2008
'Objective: This is where people select a fish tank for their pet fish.
'Other Comments:

Option Explicit

Private Sub cmd5Gall_Click(Index As Integer)
Open App.Path & "\pictanks.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Tanks(CTR)
Loop
PicResults.Picture = LoadPicture(Tanks(2))
Close #1
End Sub

Private Sub cmd10Gall_Click(Index As Integer)
PicResults.Cls
Open App.Path & "\picTanks.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Tanks(CTR)
Loop

PicResults.Picture = LoadPicture(Tanks(3))
Close #1
End Sub

Private Sub cmd2Gall_Click(Index As Integer)
PicResults.Cls
Open App.Path & "\picTanks.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Tanks(CTR)
Loop
PicResults.Picture = LoadPicture(Tanks(1))
Close #1
End Sub

Private Sub cmdMainMenu_Click()
frmMainMenu.Show
frmFishTanks.Hide
End Sub

Private Sub cmdNoThanks_Click()
frmFishTanks.Hide
frmFishAcc.Show
End Sub

Private Sub cmdFive_Click()
MsgBox ("You have purchased a five gallon fish tank for $50.00.")
HabitatCost = HabitatCost + 50
frmFishAcc.Show
frmFishTanks.Hide
End Sub

Private Sub cmdTen_Click()
MsgBox ("You have purchased a ten gallon fish tank for $75.00.")
HabitatCost = HabitatCost + 75
frmFishAcc.Show
frmFishTanks.Hide
End Sub

Private Sub cmdTwo_Click()
MsgBox ("You have purchased a two gallon fish tank for $20.00.")
HabitatCost = HabitatCost + 20
frmFishAcc.Show
frmFishTanks.Hide
End Sub


