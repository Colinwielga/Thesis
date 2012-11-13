VERSION 5.00
Begin VB.Form frmRoulette 
   Caption         =   "Roulette"
   ClientHeight    =   9420
   ClientLeft      =   2715
   ClientTop       =   1185
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   10200
   Begin VB.PictureBox picRoulette 
      Height          =   10335
      Left            =   -480
      Picture         =   "frmRoulette.frx":0000
      ScaleHeight     =   10275
      ScaleWidth      =   10635
      TabIndex        =   0
      Top             =   -840
      Width           =   10695
      Begin VB.CommandButton cmdKey 
         Height          =   495
         Left            =   7200
         Picture         =   "frmRoulette.frx":3F4EC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Would the dealer notice if you stole his keys? It's a gamble..."
         Top             =   9240
         Width           =   855
      End
      Begin VB.PictureBox picResults 
         BackColor       =   &H0080FF80&
         Height          =   735
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   5955
         TabIndex        =   5
         Top             =   960
         Width           =   6015
      End
      Begin VB.CommandButton cmdBlack 
         BackColor       =   &H00000000&
         Height          =   735
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Place bet on BLACK"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdRed 
         BackColor       =   &H000000FF&
         Height          =   735
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Place bet on RED"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdSpinWheel 
         BackColor       =   &H008080FF&
         Caption         =   "Spin"
         Height          =   1215
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8760
         Width           =   1935
      End
      Begin VB.CommandButton cmdLeave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Leave Roulette Area"
         Height          =   1215
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   8760
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8640
      Top             =   600
   End
End
Attribute VB_Name = "frmRoulette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Option Explicit
'This form allows the user to gamble on roulette
Dim color As String
Dim bet As Single

Private Sub cmdBlack_Click()
    'This button allows the user to gamble on black
    'If the user bets higher than their current balance, the user wouldn't be able to make the bet
    If balanceglobal <= 50000 Then
        bet = 0
        bet = InputBox("Please enter the amount you wish to bet:", "Bet")
        color = "BLACK"
        If bet > balanceglobal Then
            MsgBox "We aren't a charity, get more money", , "Insufficient Funds"
            bet = 0
        End If
        picResults.Cls
        picResults.Print "You bet " & FormatCurrency(bet) & " on " & color & ".  Good Luck to you."
    Else
        MsgBox "We suspect you of cheating and you need to leave the Roulette table.", , "Cheater?"
        frmRoulette.Hide
        frmLobby.Show
    End If
End Sub

Private Sub cmdKey_Click()
    'Pick up key to have ability to enter Casino Storage Room
    cmdKey.Visible = False
    frmLobby.cmdEntrance.Enabled = True
End Sub

Private Sub cmdLeave_Click()
    'Go back to Lobby
    frmRoulette.Hide
    frmLobby.Show
End Sub

Private Sub cmdRed_Click()
    'This button allows the user to gamble on red
    'If the user bets higher than their current balance, the user wouldn't be able to make the bet
    If balanceglobal <= 50000 Then
        bet = 0
        bet = InputBox("Please enter the amount you wish to bet:", "Bet")
        color = "RED"
        If bet > balanceglobal Then
            MsgBox "We aren't a charity, get more money", , "Insufficient Funds"
            bet = 0
        End If
        picResults.Cls
        picResults.Print "You bet " & FormatCurrency(bet) & " on " & color & ".  Good Luck to you."
    Else
        MsgBox "We suspect you of cheating and you need to leave the Roulette table.", , "Cheater?"
        frmRoulette.Hide
        frmLobby.Show
    End If
    
End Sub

Private Sub cmdSpinWheel_Click()
    'Enables Timer
    Timer1.Enabled = True
End Sub

Private Sub Form_Initialize()
    'Tells Rules to user and disables timer
    Timer1.Enabled = False
    MsgBox "To play Roulette choose either the color red or black", , "Rules"
    MsgBox "Place a bet and spin the wheel", , "Rules"
    MsgBox "If you chose the right color, you win", , "Rules"
End Sub

Private Sub Timer1_Timer()
    'The roulette wheel spins
    Dim counter As Integer
    For counter = 1 To 10
        picRoulette.Picture = LoadPicture(App.Path & "\roulette.jpg")
        Sleep (40)
        picRoulette.Picture = LoadPicture(App.Path & "\roulette1.5.jpg")
        Sleep (40)
        picRoulette.Picture = LoadPicture(App.Path & "\roulette2.jpg")
        Sleep (40)
        picRoulette.Picture = LoadPicture(App.Path & "\roulette2.5.jpg")
        Sleep (40)
        picRoulette.Picture = LoadPicture(App.Path & "\roulette3.jpg")
        Sleep (40)
        picRoulette.Picture = LoadPicture(App.Path & "\roulette3.5.jpg")
        Sleep (40)
        picRoulette.Picture = LoadPicture(App.Path & "\roulette4.jpg")
        Sleep (40)
        picRoulette.Picture = LoadPicture(App.Path & "\roulette4.5.jpg")
        Sleep (40)
    Next counter
    Timer1.Enabled = False
    'When the user's balance is equal to an even number, they lose
    'When the user's balance is equal to an odd number, they win
    If LCase(color) <> "red" And LCase(color) <> "black" Then
        MsgBox "Please choose a color", , "Color"
    Else
        If (balanceglobal / 2) <> Int(balanceglobal / 2) Then
            MsgBox "Congratulations " & color & " is the winning color.", , "WINNER"
            MsgBox "You win " & FormatCurrency(bet) & ".", , "Congrats"
            balanceglobal = balanceglobal + bet
            picResults.Print "*************************************************"
            picResults.Print "Your wallet now holds " & FormatCurrency(balanceglobal) & "."
        Else
            MsgBox color & " is not the winning color", , "Sorry"
            MsgBox "Thanks for playing", , "Come back soon"
            balanceglobal = balanceglobal - bet
            picResults.Print "*************************************************"
            picResults.Print "Your wallet now holds " & FormatCurrency(balanceglobal) & "."
        End If
    End If
    bet = 0
    color = ""
End Sub

