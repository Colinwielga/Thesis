VERSION 5.00
Begin VB.Form frmEnterUserStats 
   BackColor       =   &H80000018&
   Caption         =   "User Stats"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnHome 
      Caption         =   "Return to Home Screen"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3240
      TabIndex        =   14
      Top             =   7560
      Width           =   2535
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   7560
      Width           =   2535
   End
   Begin VB.CommandButton btnAvgWr 
      Caption         =   "Find Average Number of Yards (WR)"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   11
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CommandButton btnAvgRb 
      Caption         =   "Find Average Number of Yards (RB)"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   10
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CommandButton btnAvg 
      Caption         =   "Find Average Number of Yards (QB)"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   9
      Top             =   6240
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3240
      ScaleHeight     =   2715
      ScaleWidth      =   9075
      TabIndex        =   8
      Top             =   3120
      Width           =   9135
   End
   Begin VB.CommandButton btnStatsWR 
      Caption         =   " Enter WR Stats"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton btnStatsRB 
      Caption         =   "Enter RB Stats"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton btnStatsQB1 
      Caption         =   "Enter QB Stats"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtNumber 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Enter Your Own Football Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   12255
   End
   Begin VB.Label lblPosition 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "3. Enter Stats Based on Your Position"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label lblNumber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "2. Enter Number==>"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "1. Enter Name==>"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   7920
      Picture         =   "frmEnterUserStats.frx":0000
      Top             =   960
      Width           =   2160
   End
End
Attribute VB_Name = "frmEnterUserStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim playerName As String
Dim playerNumber As Single
Dim playerPos As String
Dim qbTds As Single
Dim qbAtt As Single
Dim qbComp As Single
Dim qbYards As Single
Dim rbTds As Single
Dim rbAtt As Single
Dim rbYards As Single
Dim wrRec As Single
Dim wrTds As Single
Dim wrYards As Single
Dim ctr As Integer
Private Sub btnAvg_Click() 'this button will find the avg number of yards per pass for qb's
Dim Avg As Single 'declare average variable

picResults.Cls 'clear picture box for new results

'assign variables
'user will input stats in an input box
qbComp = InputBox("Enter the number of passes you have completed.")
qbYards = InputBox("Enter the number of yards you have thrown for.")
playerPos = "Quarter Back"
Avg = qbYards / qbComp 'calcualtes average

picResults.Print "Player Name: "; playerName
picResults.Print "Number: "; playerNumber
picResults.Print "Position: "; playerPos
picResults.Print
picResults.Print "Your average yards per completed pass is "; FormatNumber(Avg, 2)
End Sub
Private Sub btnAvgRb_Click() 'this button will find the avg number of yards per run for rb's
Dim Avg As Single

picResults.Cls 'clear picture box for new results

'assign variables
'user will input stats in an input box
rbAtt = InputBox("Enter the number of runs you have attempted.")
rbYards = InputBox("Enter the number of yards you have gained.")
playerPos = "Running Back"
Avg = rbYards / rbAtt 'calcualtes average

picResults.Print "Player Name: "; playerName
picResults.Print "Number: "; playerNumber
picResults.Print "Position: "; playerPos
picResults.Print
picResults.Print "Your average yards per run is "; FormatNumber(Avg, 2)
End Sub
Private Sub btnAvgWr_Click() 'this button will find the avg number of yards per catch for wr's
Dim Avg As Single 'declare average variable

picResults.Cls 'clear picture box for new results

'assign variables
'user will input stats in an input box
wrRec = InputBox("Enter the number of passes you have caught.")
wrYards = InputBox("Enter the number of yards you have gained.")
playerPos = "Wide Receiver"
Avg = wrYards / wrRec 'calcualtes average

picResults.Print "Player Name: "; playerName
picResults.Print "Number: "; playerNumber
picResults.Print "Position: "; playerPos
picResults.Print
picResults.Print "Your average yards per catch is "; FormatNumber(Avg, 2)
End Sub
Private Sub btnClear_Click(Index As Integer) 'this button clears the picture box
picResults.Cls
End Sub
Private Sub btnHome_Click(Index As Integer) 'this button returns user back to home page
    frmEnterUserStats.Hide
    frmStart.Show
End Sub
Private Sub btnQuit_Click(Index As Integer) 'this button quits out of program
End
End Sub
Private Sub btnStatsQB1_Click(Index As Integer) 'this button will allow the user to input their qb stats in input boxes

picResults.Cls 'clear picture box for new results

'assign variables
'user will enter stats in input boxes
qbAtt = InputBox("Enter the number of passes you have attempted.")
qbComp = InputBox("Enter the number of passes you have completed.")
qbYards = InputBox("Enter the number of yards you have thrown for.")
qbTds = InputBox("Enter the number of touchdowns you have thrown for.")
playerName = txtName.Text
playerNumber = txtNumber.Text
playerPos = "Quarter Back"
picResults.Print "Player Name: "; playerName
picResults.Print "Number: "; playerNumber
picResults.Print "Position: "; playerPos
picResults.Print
picResults.Print "Att.", "Comp.", "Yards", "Touchdowns"
picResults.Print "*************************************************************************************************"
picResults.Print qbAtt, qbComp, qbYards, qbTds

End Sub
Private Sub btnStatsRB_Click()

picResults.Cls 'clear picture box for new results

'assign variables
'user will enter stats in input boxes
rbAtt = InputBox("Enter the number of runs you have attempted.")
rbYards = InputBox("Enter the number of yards you have gained.")
rbTds = InputBox("Enter the number of touchdowns you have scored.")
playerName = txtName.Text
playerNumber = txtNumber.Text
playerPos = "Running Back"

picResults.Print "Player Name: "; playerName
picResults.Print "Number: "; playerNumber
picResults.Print "Position: "; playerPos
picResults.Print
picResults.Print "Att.", "Yards", "Touchdowns"
picResults.Print "*************************************************************************************************"
picResults.Print rbAtt, rbYards, rbTds
End Sub
Private Sub btnStatsWR_Click()

picResults.Cls 'clear picture box for new results

'assign variables
'user will enter stats in input boxes
wrRec = InputBox("Enter the number of passes you have caught.")
wrYards = InputBox("Enter the number of yards you have gained.")
wrTds = InputBox("Enter the number of touchdowns you have scored.")
playerName = txtName.Text
playerNumber = txtNumber.Text
playerPos = "Wide Receiver"
picResults.Print "Player Name: "; playerName
picResults.Print "Number: "; playerNumber
picResults.Print "Position: "; playerPos
picResults.Print
picResults.Print "Rec", "Yards", "Touchdowns"
picResults.Print "*************************************************************************************************"
picResults.Print wrRec, qbYards, qbTds
End Sub

Private Sub helmet_Click()

End Sub
