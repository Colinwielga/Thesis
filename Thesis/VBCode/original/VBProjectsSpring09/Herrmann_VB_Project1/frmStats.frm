VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   4560
   ClientTop       =   2655
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   7695
   Begin VB.PictureBox picExplain 
      Height          =   735
      Left            =   4320
      ScaleHeight     =   675
      ScaleWidth      =   2355
      TabIndex        =   12
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H000080FF&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H000000FF&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.PictureBox picCalculator 
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display names and scores"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Points After"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Drop Goals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Tries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblCalc 
      Alignment       =   2  'Center
      Caption         =   "                                                                 Total Points Calculator"
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim names(1 To 100) As String, tries(1 To 100) As Integer, dropgoals(1 To 100) As Integer, pafter(1 To 100) As Integer

Private Sub cmdCalculate_Click()

Dim try As Integer, dropgoal As Integer, ptafter As Integer, total As Integer

picCalculator.Cls
picExplain.Cls

try = txt1
dropgoal = txt2
ptafter = txt3

total = try * 5 + dropgoal * 3 + ptafter * 2
picCalculator.Print total

picExplain.Print "Tries are worth 5 points each"
picExplain.Print "Drop goals are worth 3 points"
picExplain.Print "Points after tries are worth 2 pts"

If total > 99 Then
    MsgBox "That total can't even fit!", , "Wow"
ElseIf total < 0 Then
    MsgBox "Re-enter values using positive integers", , "Uh-Oh"
End If

End Sub

Private Sub cmdDisplay_Click()

picResults.Cls
Open App.Path & "\statsTrys.txt" For Input As #1

picResults.Print "Name"; Tab(22); "Tries"; Tab(30); "DropGoals"; Tab(43); "Pt. After"
picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Do While Not EOF(1)
    CTR = CTR + 1
        Input #1, names(CTR), tries(CTR), dropgoals(CTR), pafter(CTR)
        picResults.Print names(CTR); Tab(23); tries(CTR); Tab(33); dropgoals(CTR); Tab(45); pafter(CTR)
Loop
Close #1

End Sub

Private Sub cmdMenu_Click()
frmMenu.Show
frmStats.Hide
End Sub

