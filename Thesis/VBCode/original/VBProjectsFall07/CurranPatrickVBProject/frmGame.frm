VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Video Poker"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   Picture         =   "frmGame.frx":0000
   ScaleHeight     =   7035
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCashOut 
      Caption         =   "Cash Out"
      Height          =   1455
      Left            =   4680
      TabIndex        =   10
      Top             =   5520
      Width           =   1455
   End
   Begin VB.PictureBox picCredits 
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   5520
      Width           =   2055
   End
   Begin VB.PictureBox picWinnings 
      Height          =   615
      Left            =   6840
      ScaleHeight     =   555
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdDeal 
      Caption         =   "Deal!"
      Height          =   1455
      Left            =   3120
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.PictureBox picCard5 
      BackColor       =   &H80000007&
      Height          =   2175
      Left            =   7080
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox picCard4 
      BackColor       =   &H80000007&
      Height          =   2175
      Left            =   5400
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox picCard3 
      BackColor       =   &H80000007&
      Height          =   2175
      Left            =   3840
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox picCard2 
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   2160
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox picCard1 
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   600
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "        Credits"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "     Winnings"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   6360
      Width           =   1815
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCashOut_Click()
MsgBox "Cashing Out: " & FormatCurrency(Credits)
End
End Sub

Private Sub cmdDeal_Click()
Dim Pos As Integer     'Declares Variables in their proper format.
Dim X As Integer
Dim Y As Integer
Dim Temp As String
Dim Hand As String
Dim answer As Boolean

answer = False    'Sets variable to false.

Randomize       'Initiates randomization function.

Credits = Credits - 5   'Subtracts 5 credits for every hand played.

If Credits <= 0 Then    'Determines whether or not player has enough money to play.

    MsgBox ("You are broke!")
    End
End If

picCredits.Cls      'Clears the Credit picturebox for the next hand.
picWinnings.Cls     'Clears the Winnings picturebox for the next hand.
picCredits.Print Credits    'Displays new amount of Credits.
picCard1.Cls
picCard2.Cls
picCard3.Cls        'Clears the Cards for the next hand.
picCard4.Cls
picCard5.Cls

'Selects two random positions in the Cards array and swaps them.
    'It repeats this process 500 times in order to shuffle the cards in the array.
For Pos = 1 To 500
    X = Int(Rnd * 52) + 1
    Y = Int(Rnd * 52) + 1
    Temp = Cards(X)
    Cards(X) = Cards(Y)
    Cards(Y) = Temp
Next Pos

'Displays the first 5 cards of the Cards array, hence the first 5 cards off of the top of the deck.

picCard1.Picture = LoadPicture(App.Path & "\" & Cards(1))
picCard2.Picture = LoadPicture(App.Path & "\" & Cards(2))
picCard3.Picture = LoadPicture(App.Path & "\" & Cards(3))
picCard4.Picture = LoadPicture(App.Path & "\" & Cards(4))
picCard5.Picture = LoadPicture(App.Path & "\" & Cards(5))

'Determines whether or not the user has entered a valid poker hand.
'If a winning hand is entered, it computes the amount of credits the player receives.
'If an invalid hand is entered, then it displays an error and asks the user to input a valid hand.
Do Until answer = True
    Hand = InputBox("What hand did you receive? ex.(none,pair,3 of a kind,flush,etc.) please use lower case letters.")   'Asks the user to input the name of the poker hand they received.
    
    If Hand = "none" Then
            picWinnings.Print "You Lose"
            answer = True       'Causes the program to break out of the loop.
        ElseIf Hand = "pair" Then
            Credits = Credits + 5
            picWinnings.Print FormatCurrency(5)
            answer = True
        ElseIf Hand = "two pair" Then
            Credits = Credits + 5
            picWinnings.Print FormatCurrency(5)
            answer = True
        ElseIf Hand = "3 of a kind" Or Hand = "three of a kind" Then
            Credits = Credits + 15
            picWinnings.Print FormatCurrency(15)
            answer = True
        ElseIf Hand = "straight" Then
            Credits = Credits + 25
            picWinnings.Print FormatCurrency(25)
            answer = True
        ElseIf Hand = "flush" Then
            Credits = Credits + 35
            picWinnings.Print FormatCurrency(35)
            answer = True
        ElseIf Hand = "full house" Then
            Credits = Credits + 45
            picWinnings.Print FormatCurrency(45)
            answer = True
        ElseIf Hand = "4 of a kind" Or Hand = "four of a kind" Then
            Credits = Credits + 250
            picWinnings.Print FormatCurrency(250)
            answer = True
        ElseIf Hand = "straight flush" Then
            Credits = Credits + 250
            picWinnings.Print FormatCurrency(250)
            answer = True
        ElseIf Hand = "royal flush" Then
            Credits = Credits + 4000
            picWinnings.Print FormatCurrency(4000)
            answer = True
        Else
            MsgBox ("Invalid hand") 'Informs the player if they have entered an invalid hand.
            answer = False      'Causes the program to loop since answer is not equal to true.
    End If
Loop    'Loops back to the beginning if player has entered an invalid hand.


End Sub
