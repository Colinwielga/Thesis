VERSION 5.00
Begin VB.Form frmLigue1 
   BackColor       =   &H000080FF&
   Caption         =   "Ligue 1 (France)"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1800
      Left            =   4320
      Picture         =   "Ligue1.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   13
      Top             =   4080
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1365
      Left            =   1200
      Picture         =   "Ligue1.frx":9DF2
      ScaleHeight     =   1305
      ScaleWidth      =   1965
      TabIndex        =   12
      Top             =   4320
      Width           =   2025
   End
   Begin VB.CommandButton cmdGetResult 
      BackColor       =   &H00FF0000&
      Caption         =   "Get the result of the game"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdBet 
      BackColor       =   &H00FF0000&
      Caption         =   "Place a Bet"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF0000&
      Caption         =   "Exit"
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox lstGames 
      Height          =   1185
      ItemData        =   "Ligue1.frx":124C8
      Left            =   240
      List            =   "Ligue1.frx":124CA
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton cmdCashOut 
      BackColor       =   &H00FF0000&
      Caption         =   "Cash Out"
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturnSoccer 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to Soccer Homepage"
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturnHomepage 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to Homepage"
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FF0000&
      Caption         =   "Display Games"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblBetCTR 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblNumBets 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Caption         =   "# of bets:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblTotalAmount 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblAccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Caption         =   "Account:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmLigue1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sports Betting Project
'frmLigue1 (French soccer)
'Written by: Sean Egan
'Written on: 3/21/09
'This form allows the user to bet on French soccer games.

'Declare variables
Dim HomeTeam(1 To 50) As String, AwayTeam(1 To 50) As String
Dim r As Integer
Dim Team As String

Private Sub cmdBet_Click()
    'This button allows the user to place a bet. It decides whether
    ' the user has checked a box. If they have, it prompts them
    ' to choose a team to bet on and the amount. If they don't
    ' have enough money, it opens a message box saying so.
    
    'Declare the variable
    Dim j As Integer
    
    For j = 0 To lstGames.ListCount - 1
        If lstGames.Selected(j) = True Then
            'Prompts the user to choose a team to bet on
            Team = InputBox("Which team would you like to bet on?")
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
            'Message box with the team and amount they bet
            MsgBox ("You bet " & FormatCurrency(Bet) & " on " & UCase$(Team) & ".")
        End If
    Next j
    
    'Enable Result button
    cmdGetResult.Enabled = True
    
End Sub

Private Sub cmdCashOut_Click()
    'Loads the Cash Out form
    frmCashOut.Show
    'Hides the Ligue1 form
    frmLigue1.Hide
End Sub

Private Sub cmdDisplay_Click()
    'This button reads all of the teams from a file into 2 parallel
    ' arrays and prints them in a listbox randomly
    
    'Declare variables
    Dim Num(1 To 50) As Integer
    Dim CTR As Integer, rCTR As Integer
    Dim mylist As String
    'Set counter at 0
    CTR = 0

    'Open the file
    Open App.Path & "\Ligue1Games.txt" For Input As #1
    
    'Read the file into 2 parallel arrays
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, HomeTeam(CTR), AwayTeam(CTR)
    Loop
    'Close the file
    Close #1
    
    'Randomly selects a game from the file and displays it into the
    ' listbox. It will not print a game multiple times.
    'Set the random counter at 0
    rCTR = 0
    Do While rCTR < 10
        
        r = Int(10 * Rnd + 1)

        If Num(r) = 0 Then
            rCTR = rCTR + 1
            Num(r) = r
            lstGames.AddItem HomeTeam(r) & " vs. " & AwayTeam(r) & vbCrLf
        End If
    Loop
    
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
        MsgBox ("Congratulations! " & UCase$(Team) & " won! You win " & FormatCurrency(Bet * 2))
        'Doubles the users bet and adds it to their account
        Total = Total + (Bet * 2)
    'Code executed if the number is not a 1 (if it's a 0)
    Else
        'Message saying they have lost the bet
        MsgBox ("Sorry. " & UCase$(Team) & " lost. Better luck next time.")
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
        'Hides Ligue1 form
        frmLigue1.Hide
    End If
    
End Sub

Private Sub cmdReturnHomepage_Click()
    'Loads Homepage
    frmHomepage.Show
    'Hides Ligue1 form
    frmLigue1.Hide
End Sub

Private Sub cmdReturnSoccer_Click()
    'Loads Soccer form
    frmSoccer.Show
    'Hides Ligue1 form
    frmLigue1.Hide
End Sub

Private Sub Form_Load()
    'Loads Randomize function
    Randomize
    'Carries the total over from the previous form
    Total = Total
    'Prints the current amount in the account
    lblTotalAmount.Caption = FormatCurrency(Total)
    'Prints the current bet count
    lblBetCTR.Caption = BetCTR
End Sub
