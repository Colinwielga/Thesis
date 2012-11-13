VERSION 5.00
Begin VB.Form formGame 
   BackColor       =   &H80000008&
   Caption         =   "Game"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBalance 
      Height          =   975
      Left            =   7080
      ScaleHeight     =   915
      ScaleWidth      =   2475
      TabIndex        =   10
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdletters 
      Caption         =   "Show Letters "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtCon 
      Height          =   735
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdSolve 
      Caption         =   "Solve Puzzle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   6
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy A Vowel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   5
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdSpin 
      BackColor       =   &H8000000D&
      Caption         =   "Spin Wheel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   4
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox txtNum 
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton CmdShow 
      BackColor       =   &H80000013&
      Caption         =   "Show Puzzle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.PictureBox picPuzzle 
      Height          =   4095
      Left            =   240
      ScaleHeight     =   4035
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   2400
      Width           =   7575
   End
   Begin VB.Image vanna 
      Height          =   3855
      Left            =   8280
      Picture         =   "formGame.frx":0000
      Top             =   2520
      Width           =   2700
   End
   Begin VB.Image Imagemoney 
      Height          =   1470
      Left            =   10080
      Picture         =   "formGame.frx":4BE1
      Top             =   240
      Width           =   1290
   End
   Begin VB.Label lblCash 
      BackColor       =   &H8000000D&
      Caption         =   "Total Cash Won"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblNum 
      BackColor       =   &H8000000D&
      Caption         =   "Enter # 1 to 3 to Choose Puzzle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblCon 
      BackColor       =   &H8000000D&
      Caption         =   "Input Constanant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "formGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Wheeltime.vbp
'Form Name:formGame.frm
'Author: Kristen Palomo
'Date written: October 31, 2005
'Objective: The code written allows a user to play my version of Wheel of Fortune


Dim X As Double
Dim SpacesArray(1 To 3) As String
Dim Prizes(1 To 10) As Single
Dim PrizesWon As Single
Dim winnings As Single
Dim Con As String
Dim Letters(1 To 50) As String
Dim I As Integer, J As Integer

Private Sub cmdBuy_Click()
Dim Vowel As String
Vowel = InputBox("A Vowel Costs $200 Please Enter Vowel")
winnings = winnings - 200
picBalance.Cls
picBalance.Print FormatCurrency(winnings)
txtCon.Text = Vowel
cmdletters_Click


End Sub

Private Sub cmdletters_Click()

picPuzzle.Cls
Con = txtCon.Text
Open App.Path & "\letters" & X & ".txt" For Input As #1
I = 0
Do Until EOF(1)
I = I + 1
    Input #1, Letters(I)
Loop
picPuzzle.FontSize = 30
picPuzzle.Print SpacesArray(X)
For J = 1 To I
 If Con = Letters(J) Then
    picPuzzle.Print Con;
    winnings = winnings + PrizesWon
    picBalance.Cls
    picBalance.Print FormatCurrency(winnings)
 ElseIf Not (Letters(J) = " ") Then
    picPuzzle.Print " ";
 Else
    picPuzzle.Print "_";
 End If
 Next J
picPuzzle.Print
Close #1
 
End Sub

Private Sub CmdShow_Click()

Dim I As Integer
X = Val(txtNum.Text)
Open App.Path & "\puzzles.txt" For Input As #1
For I = 1 To 3
    Input #1, SpacesArray(I)
    
Next I
If X = 1 Then
    MsgBox ("Quote"), , "Old School"
    ElseIf X = 2 Then
    MsgBox ("Proper Name"), , "Old School"
    Else
    MsgBox ("Place"), , "Old School"
    End If
    
Close #1
picPuzzle.FontSize = 30
picPuzzle.Print SpacesArray(X)
    
End Sub

Private Sub cmdSolve_Click()
    Dim Choices(1 To 3) As String
    Dim I, J As Integer
    Dim NotFound As Boolean
    Open App.Path & "\choices.txt" For Input As #1
  
    
    Answers = InputBox("Type in Answer", "Solve Puzzle")
    For I = 1 To 3
        Input #1, Choices(I)
    Next I
    I = I - 1
    J = 0
    NotFound = True
    Do While J < I And NotFound
        J = J + 1
        If Answers = Choices(J) Then
            NotFound = False
        End If
    Loop
        
    If Not (NotFound) Then
        MsgBox "Congratulations!", , "You Won"
        frmStart.Visible = True
        formGame.Visible = False
    Else
        MsgBox ("Sorry Game Over"), , "Error"
        frmStart.Visible = True
        formGame.Visible = False
    End If
    Close #1
End Sub
Private Sub cmdSpin_Click()

Dim I As Integer
Open App.Path & "\prizes.txt" For Input As #1
    For I = 1 To 10
        Input #1, Prizes(I)
    Next I
Close #1
rand = Int(10 * Rnd) + 1
PrizesWon = Prizes(rand)
If Not PrizesWon = 0 Then
MsgBox FormatCurrency(PrizesWon), , "Amount"
Else
picBalance.Cls
winnings = 0
MsgBox "You went Bankrupt.", , "Bankrupt"
End If

End Sub

