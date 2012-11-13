VERSION 5.00
Begin VB.Form frmPlay 
   BackColor       =   &H008080FF&
   Caption         =   "Play Bingo!!"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form2"
   ScaleHeight     =   7065
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBingo 
      Height          =   6975
      Left            =   0
      Picture         =   "VB Project Form 2.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   7515
      TabIndex        =   3
      Top             =   0
      Width           =   7575
   End
   Begin VB.CommandButton cmdCallNextNumber 
      Caption         =   "Call Next Number"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdTransferForm 
      Caption         =   "Go Back to Bingo Card"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BingoProject
'Play Form
'Missy Ulrich & Beckie Hoyt
'November 3, 2006
'This form allows the user to play the game of bingo when the program chooses a random number between 1 and 75 and checks the card on the previous
'Create form to see if the random number matches any of the numbers located on the bingo card that has either been made or received from the text file
'If the random number matches one of the numbers on the user's bingo card then an "X" is placed where the number in that textbox location used to be
'The user also has the option of returning to the previous form to check for a bingo on their card, or to exit the program altogether

Option Explicit
Dim HighValue As Integer
Dim LowValue As Integer
Dim RandomInteger As Integer
'Declares the variables used in the form

Private Sub cmdCallNextNumber_Click()
    Size = 5
    Randomize
    HighValue = 75
    LowValue = 1
    'Sets the variables to specific values
    RandomInteger = Int((HighValue - LowValue + 1) * Rnd + LowValue)
    MsgBox RandomInteger
    'This formula picks a random number between 1 and 75 and displays that number to the user in a messagebox on the screen
    Dim Found As Boolean
    Found = False
    'Initializes the search to show that the number being searched for has not been found yet
    For Pos = 1 To Size
        If RandomInteger = BArray(Pos) Then
            frmCreate.txtB(Pos).Text = "X"
            Found = True
            'This loops through the B array to see if the random number selected by the program matches a number in the B array textboxes
            'If the numbers match, then an "X" is displayed to the user in the place of the old number in the textbox
        ElseIf RandomInteger = IArray(Pos) Then
            frmCreate.txtI(Pos).Text = "X"
            Found = True
            'This loops through the I array to see if the random number selected by the program matches a number in the I array textboxes
            'If the numbers match, then an "X" is displayed to the user in the place of the old number in the textbox
        ElseIf RandomInteger = NArray(Pos) Then
            If Pos < 5 Then
                frmCreate.txtN(Pos).Text = "X"
                Found = True
            End If
            'This loops through the N array to see if the random number selected by the program matches a number in the N array textboxes
            'If the numbers match, then an "X" is displayed to the user in the place of the old number in the textbox
        ElseIf RandomInteger = GArray(Pos) Then
             frmCreate.txtG(Pos).Text = "X"
             Found = True
             'This loops through the G array to see if the random number selected by the program matches a number in the G array textboxes
             'If the numbers match, then an "X" is displayed to the user in the place of the old number in the textbox
        ElseIf RandomInteger = OArray(Pos) Then
             frmCreate.txtO(Pos).Text = "X"
             Found = True
            'This loops through the O array to see if the random number selected by the program matches a number in the O array textboxes
            'If the numbers match, then an "X" is displayed to the user in the place of the old number in the textbox
        End If
    Next Pos
    'Once all of the arrays have been searched through to see if the numbers match, and if the number on the bingo card didn't match the random number, then the loop ends and the search is still seen as not found
    If Found = False Then
            MsgBox "Sorry, Your Card Does Not Have That Number.", , "Error"
    End If
    'If the the random number could not be matched to a corresponding number on the user's bingo card from the Create form, then the above messagebox displays to the user
    'that the number was not on their bingo card
End Sub

Private Sub cmdQuit_Click()
    End
    'The command button allows the user to exit the program
End Sub

Private Sub cmdTransferForm_Click()
    frmCreate.Show
    frmPlay.Hide
    'This allows the user to switch from the Play form and go back to the Create form to check for a bingo
End Sub

