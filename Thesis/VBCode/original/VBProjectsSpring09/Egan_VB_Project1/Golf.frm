VERSION 5.00
Begin VB.Form frmGolf 
   BackColor       =   &H00004000&
   Caption         =   "Golf"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   4320
      Picture         =   "Golf.frx":0000
      ScaleHeight     =   1485
      ScaleWidth      =   1545
      TabIndex        =   13
      Top             =   4320
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1950
      Left            =   1200
      Picture         =   "Golf.frx":78EA
      ScaleHeight     =   1890
      ScaleWidth      =   1575
      TabIndex        =   12
      Top             =   3960
      Width           =   1635
   End
   Begin VB.CommandButton cmdCashOut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cash Out"
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetResult 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get the result of the match"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Place Bet"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtGolfers 
      Height          =   1935
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display Golfers"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Homepage"
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblBetCTR 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblNumBets 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      Caption         =   "# of bets:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblTotalAmount 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblAccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      Caption         =   "Account:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblWhichGolfer 
      BackColor       =   &H00004000&
      Caption         =   "Which golfer do you want to bet on to win the Masters?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmGolf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sports Betting Project
'frmGolf
'Written by: Sean Egan
'Written on: 3/19/09
'This form allows the user to bet on who will win the upcoming
' Masters golf tournament.

'Declare variables
Dim Golfer As String, Bet As Single
Dim CTR As Integer

Private Sub cmdBet_Click()
    'This button allows the user to place a bet. It prompts them
    ' to choose a team to bet on and the amount. If they don't
    ' have enough money, it opens a message box saying so.
    
    
        'Prompts the user to choose a golfer to bet on
        Golfer = InputBox("Which golfer would you like to bet on?")
        'Prompts the user to choose an amount to bet with
        Bet = InputBox("How much would you like to bet?")
        'If the bet amount exceeds their account, a message is
        ' returned saying so.
        If Bet > Total Then
            Bet = 0
            MsgBox ("I'm sorry. You have insufficient funds.")
            'Prompts the user to choose a different amount.
            Bet = InputBox("How much would you like to bet?")
        End If
        'Message box with the golfer and amount they bet
        MsgBox ("You bet " & FormatCurrency(Bet) & " on " & UCase$(Golfer) & ".")
    
    'Enable Result button
    cmdGetResult.Enabled = True
End Sub

Private Sub cmdCashOut_Click()
    'Loads Cash Out form
    frmCashOut.Show
    'Hides Golf form
    frmGolf.Hide
End Sub

Private Sub cmdDisplay_Click()
    'This button reads the golfers into an array and prints them
    ' in a text box with a scroll bar
    
    'Declare the variables
    Dim Golfers(1 To 50) As String
    Dim Num(1 To 50) As Integer
    Dim mytext As String
    
    'Set the counter to 0
    CTR = 0
    
    'Open the file
    Open App.Path & "\Golfers.txt" For Input As #1
    
    'Read the golfers into an array until the end of the file
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Golfers(CTR)
        'Print the names into the text box
        mytext = mytext & Golfers(CTR) & vbCrLf
        txtGolfers.Text = mytext
    Loop
    
    'Close the file
    Close #1
    
    'Enable Bet button
    cmdBet.Enabled = True
    
End Sub


Private Sub cmdExit_Click()
    'Closes the program
    End
End Sub

Private Sub cmdGetResult_Click()
    'This program decides whether the person wins or loses the bet
    
    'Declare variables
    Dim WinLose As Integer
    Dim NewTotal As Single
    
    'Randomly selects a 0 or 1
    WinLose = 1 * Rnd + 0
    
    'Code executed if the number is a 1
    If WinLose = 1 Then
        'Message telling the user they have won
        MsgBox ("Congratulations! " & UCase$(Golfer) & " won! You win " & FormatCurrency(Bet * 2))
        'Doubles the users bet and adds it to their account
        Total = Total + (Bet * 2)
    'Code executed if the number is not a 1 (if it's a 0)
    Else
        'Message saying they have lost the bet
        MsgBox ("Sorry. " & UCase$(Golfer) & " lost. Better luck next time.")
        'Subtracts the bet from the user's account
        Total = Total - Bet
    End If
    
    'Prints the new amount in the account
    lblTotalAmount.Caption = FormatCurrency(Total)
    'Increments the bet counter by 1
    BetCTR = BetCTR + 1
    'Prints the new bet count
    lblBetCTR.Caption = BetCTR
    
    'If the account has $0, it sends the user to the Out of Money form
    If Total = 0 Then
        'Loads Out of Money form
        frmOutOfMoney.Show
        'Hides Golf form
        frmGolf.Hide
    End If
End Sub

Private Sub cmdReturn_Click()
    'Loads Homepage
    frmHomepage.Show
    'Hides Golf form
    frmGolf.Hide
End Sub

Private Sub Form_Load()
    'Carries the total over from the previous form
    Total = Total
    'Prints the current amount in the account
    lblTotalAmount.Caption = FormatCurrency(Total)
    'Prints the current bet count
    lblBetCTR.Caption = BetCTR
End Sub
