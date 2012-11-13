VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000D&
   Caption         =   "Form4"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form4"
   ScaleHeight     =   8640
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit2 
      Caption         =   "Back to selection screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      TabIndex        =   21
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdpunch20 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Picture         =   "Project4.frx":0000
      TabIndex        =   20
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch19 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Picture         =   "Project4.frx":0464
      TabIndex        =   19
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch18 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Picture         =   "Project4.frx":08C8
      TabIndex        =   18
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch17 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "Project4.frx":0D2C
      TabIndex        =   17
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch16 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Picture         =   "Project4.frx":1190
      TabIndex        =   16
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch5 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Picture         =   "Project4.frx":15F4
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch10 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Picture         =   "Project4.frx":1A58
      TabIndex        =   14
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch15 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Picture         =   "Project4.frx":1EBC
      TabIndex        =   13
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdPunchaBunch 
      Caption         =   "Click Here to Start Punch a Bunch!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   12
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CommandButton cmdpunch12 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "Project4.frx":2320
      TabIndex        =   11
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch13 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Picture         =   "Project4.frx":2784
      TabIndex        =   10
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch14 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Picture         =   "Project4.frx":2BE8
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch2 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "Project4.frx":304C
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch3 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Picture         =   "Project4.frx":34B0
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch4 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Picture         =   "Project4.frx":3914
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch8 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Picture         =   "Project4.frx":3D78
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch9 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Picture         =   "Project4.frx":41DC
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch11 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Picture         =   "Project4.frx":4640
      TabIndex        =   3
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch7 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "Project4.frx":4AA4
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch6 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Picture         =   "Project4.frx":4F08
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdpunch1 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Picture         =   "Project4.frx":536C
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'In this game, the user earns chances to "punch" the board by guessing if the actual price of a presented product is higher or lower than the price shown with it in an input box.
Dim ctr As Double   'They can win up to three chances.  With each chance they can punch one of the 20 buttons on the board.  Each button is worth between $5 and $100.
                    'Based on the amount of a punch and how many chances the player has left he or she must decide to either stay with the dollar amount of the last punch or to try for a higher amount.
                    'If a player uses all of the chances earned, the dollar amount they "win" is whatever the last punch was.
                    'The 20 cmdpunch buttons all do the exact same thing.
                    'Under Option Explicit I declared my counter variable to be used by 21 of the 22 bottons on the form.

Private Sub cmdpunch1_Click()   'The cmdpunch buttons present the user with a dollar amount between $5 and $100.  The user must decide if they want to keep the given amount or continuing punching for a higher amount.
ctr = ctr - 1   'This is subtracting from the counter amount established in the first part of the game.
Dim punch1 As String
punch1 = InputBox("This punch was worth $5.  Do you want to take the $5 or try again? Enter either 'stay' or 'continue'.")
 If ctr = 0 Then
    MsgBox "That was your final punch.  You win $5."
    End 'If this was the players final punch, the game ends and the player automatically wins the amount of that last punch.
ElseIf punch1 = "stay" Then 'If the player types "stay" into the input box, they win the amount of that punch.
    MsgBox "Contratulations, you have won $5."
    End     'Choosing to stay ends the game.
ElseIf punch1 = "continue" Then 'If the player types "continue" into the input box and they have not run our of punches, they may punch the board again.
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"  'If the player enters anything other than "stay" or "continue" They recieve a message saying "Invalid Entry."
    End If
cmdpunch1.Visible = False   'This causes a button to disappear after it is punched.
    
End Sub

Private Sub cmdpunch10_Click()
ctr = ctr - 1
Dim punch10 As String
punch10 = InputBox("This punch was worth $100, the highest amount on the board.  Do you want to take the $100 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $100."
    End
ElseIf punch10 = "stay" Then
    MsgBox "Contratulations, you have won $100."
    End
ElseIf punch10 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch10.Visible = False

End Sub

Private Sub cmdpunch11_Click()
ctr = ctr - 1
Dim punch11 As String
punch11 = InputBox("This punch was worth $50.  Do you want to take the $50 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $50."
    End
ElseIf punch11 = "stay" Then
    MsgBox "Contratulations, you have won $50."
    End
ElseIf punch11 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch11.Visible = False

End Sub

Private Sub cmdpunch12_Click()
ctr = ctr - 1
Dim punch12 As String
punch12 = InputBox("This punch was worth $5.  Do you want to take the $5 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $5."
    End
ElseIf punch12 = "stay" Then
    MsgBox "Contratulations, you have won $5."
    End
ElseIf punch12 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch12.Visible = False

End Sub

Private Sub cmdpunch13_Click()
ctr = ctr - 1
Dim punch13 As String
punch13 = InputBox("This punch was worth $25.  Do you want to take the $25 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $25."
    End
ElseIf punch13 = "stay" Then
    MsgBox "Contratulations, you have won $25."
    End
ElseIf punch13 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch13.Visible = False

End Sub

Private Sub cmdpunch14_Click()
ctr = ctr - 1
Dim punch14 As String
punch14 = InputBox("This punch was worth $10.  Do you want to take the $10 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $10."
    End
ElseIf punch14 = "stay" Then
    MsgBox "Contratulations, you have won $10."
    End
ElseIf punch14 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch14.Visible = False
End Sub

Private Sub cmdpunch15_Click()
ctr = ctr - 1
Dim punch15 As String
punch15 = InputBox("This punch was worth $10.  Do you want to take the $10 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $10."
    End
ElseIf punch15 = "stay" Then
    MsgBox "Contratulations, you have won $10."
    End
ElseIf punch15 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch15.Visible = False
End Sub

Private Sub cmdpunch16_Click()
ctr = ctr - 1
Dim punch16 As String
punch16 = InputBox("This punch was worth $5.  Do you want to take the $5 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $5."
    End
ElseIf punch16 = "stay" Then
    MsgBox "Contratulations, you have won $5."
    End
ElseIf punch16 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch16.Visible = False
End Sub

Private Sub cmdpunch17_Click()
ctr = ctr - 1
Dim punch17 As String
punch17 = InputBox("This punch was worth $10.  Do you want to take the $10 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $10."
    End
ElseIf punch17 = "stay" Then
    MsgBox "Contratulations, you have won $10."
    End
ElseIf punch17 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch17.Visible = False
End Sub

Private Sub cmdpunch18_Click()
ctr = ctr - 1
Dim punch18 As String
punch18 = InputBox("This punch was worth $25.  Do you want to take the $25 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $25."
    End
ElseIf punch18 = "stay" Then
    MsgBox "Contratulations, you have won $25."
    End
ElseIf punch18 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch18.Visible = False
End Sub

Private Sub cmdpunch19_Click()
ctr = ctr - 1
Dim punch19 As String
punch19 = InputBox("This punch was worth $5.  Do you want to take the $5 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $5."
    End
ElseIf punch19 = "stay" Then
    MsgBox "Contratulations, you have won $5."
    End
ElseIf punch19 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch19.Visible = False
End Sub

Private Sub cmdpunch2_Click()
ctr = ctr - 1
Dim punch2 As String
punch2 = InputBox("This punch was worth $50.  Do you want to take the $50 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $50."
    End
ElseIf punch2 = "stay" Then
    MsgBox "Contratulations, you have won $50."
    End
ElseIf punch2 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch2.Visible = False
End Sub


Private Sub cmdpunch20_Click()
ctr = ctr - 1
Dim punch20 As String
punch20 = InputBox("This punch was worth $10.  Do you want to take the $10 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $10."
    End
ElseIf punch20 = "stay" Then
    MsgBox "Contratulations, you have won $10."
    End
ElseIf punch20 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch20.Visible = False
End Sub

Private Sub cmdpunch3_Click()
ctr = ctr - 1
Dim punch3 As String
punch3 = InputBox("This punch was worth $10.  Do you want to take the $10 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $10."
    End
ElseIf punch3 = "stay" Then
    MsgBox "Contratulations, you have won $10."
    End
ElseIf punch3 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch3.Visible = False
End Sub

Private Sub cmdpunch4_Click()
ctr = ctr - 1
Dim punch4 As String
punch4 = InputBox("This punch was worth $5.  Do you want to take the $5 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $5."
    End
ElseIf punch4 = "stay" Then
    MsgBox "Contratulations, you have won $5."
    End
ElseIf punch4 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch4.Visible = False
End Sub

Private Sub cmdpunch5_Click()
ctr = ctr - 1
Dim punch5 As String
punch5 = InputBox("This punch was worth $5.  Do you want to take the $5 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $5."
    End
ElseIf punch5 = "stay" Then
    MsgBox "Contratulations, you have won $5."
    End
ElseIf punch5 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch5.Visible = False
End Sub

Private Sub cmdpunch6_Click()
ctr = ctr - 1
Dim punch6 As String
punch6 = InputBox("This punch was worth $5.  Do you want to take the $5 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $5."
    End
ElseIf punch6 = "stay" Then
    MsgBox "Contratulations, you have won $5."
    End
ElseIf punch6 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch6.Visible = False
End Sub

Private Sub cmdpunch7_Click()
ctr = ctr - 1
Dim punch7 As String
punch7 = InputBox("This punch was worth $25.  Do you want to take the $25 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $25."
    End
ElseIf punch7 = "stay" Then
    MsgBox "Contratulations, you have won $25."
    End
ElseIf punch7 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch7.Visible = False
End Sub

Private Sub cmdpunch8_Click()
ctr = ctr - 1
Dim punch8 As String
punch8 = InputBox("This punch was worth $10.  Do you want to take the $10 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $10."
    End
ElseIf punch8 = "stay" Then
    MsgBox "Contratulations, you have won $10."
    End
ElseIf punch8 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch8.Visible = False
End Sub

Private Sub cmdpunch9_Click()
ctr = ctr - 1
Dim punch9 As String
punch9 = InputBox("This punch was worth $5.  Do you want to take the $5 or try again? Enter either 'stay' or 'continue'.")
If ctr = 0 Then
    MsgBox "That was your final punch.  You win $5."
    End
If punch9 = "stay" Then
    MsgBox "Contratulations, you have won $5."
    End
ElseIf punch9 = "continue" Then
    MsgBox "Punch again."
Else
    MsgBox "Invalid Entry"
    End If
cmdpunch9.Visible = False
End Sub

Private Sub cmdPunchaBunch_Click()  'In this part of the game, the player is presented with three different products each in an input box along with a false price for the product.  They must type "higher" or "lower" into the input box to indicate what they think the actual price is.
MsgBox "You will be shown three prizes along with an incorrect price for each.  You must type the words higher or lower to indicate what you think the correct price is in comparison to the given one.  Each correct answer will earn you a chance to punch the board."

Dim guess As String
Dim guess2 As String
Dim guess3 As String

ctr = 0     'Keeps track of how many correct guesses are made.
guess = InputBox("The first product is a pencil and the given price is $10.  Do you think the actual retail price is higher or lower?")
If guess = "lower" Then
    MsgBox "You are correct!"
    ctr = ctr + 1
ElseIf guess = "higher" Then
        MsgBox "I'm sorry, that is incorrect."
Else
    MsgBox "Invalid Entry"
    End If
guess2 = InputBox("The second product is a St. John's football jersey and the given price is $2.  Do you think the actual retail price is higher or lower?")
If guess2 = "higher" Then
    MsgBox "You are correct!"
    ctr = ctr + 1
ElseIf guess2 = "lower" Then
    MsgBox "I'm sorry, that is incorrect."
Else
    MsgBox "Invalid Entry"
    End If
guess3 = InputBox("The third and final product is a tube of tooth paste and the given price is $25.  Do you think the actual retail price is higher or lower?")
If guess3 = "lower" Then
MsgBox "You are correct!"
    ctr = ctr + 1
ElseIf guess3 = "higher" Then
    MsgBox "I'm sorry, that is incorrect."
Else
    MsgBox "Invalid Entry"
    End If
MsgBox ("Congratulations, you have earned " & ctr & " chances to punch the board.  Each punch will be worth a dollar amount and you must decide if you want to stay with that dollar amount or continue using your punches.  There are eight punches that are worth $5, six that are worth $10, three that are worth $25, two that are worth $50, and 1 that is worth $100.  You win whatever the value of your final punch is.")

End Sub
  


Private Sub cmdQuit2_Click() 'Allows the player to go back to Form2 and choose a different game or quit.
Form2.Show
Form4.Hide
End Sub
