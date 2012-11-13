VERSION 5.00
Begin VB.Form RachelHaney6 
   BackColor       =   &H0000FFFF&
   Caption         =   "RachelHaney6"
   ClientHeight    =   5535
   ClientLeft      =   2955
   ClientTop       =   2265
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7710
   Visible         =   0   'False
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      Height          =   4695
      Left            =   1320
      ScaleHeight     =   4635
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   720
      Width           =   6255
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display your results."
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblEnd 
      BackColor       =   &H00FF80FF&
      Caption         =   "Here is your vacation!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "RachelHaney6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RachelHaney6 (RachelHaneyVBProject5.frm)
'Rachel Haney 3/11/04
'This form will display all of the choices that the people
'chose for their vacation and how much it cost.

Private Sub cmdContinue_Click()
    RachelHaney6.Visible = False
    RachelHaney7.Visible = True
    RachelHaney6.cmdContinue.Visible = False
End Sub

Private Sub cmdDisplay_Click()
'Tells what budget and the number of people the user chose for the trip.
    picResults.Print "Here are your vacation plans."
    picResults.Print
    picResults.Print "You decided to have a budget of "; FormatCurrency(Spend); " to spend."
    picResults.Print "You are taking "; People; " people on your vacation."
    cmdTotal.Visible = True
    cmdDisplay.Visible = False

'Tells which destination the person chose.
    picResults.Print
    If City = 1 Then
            picResults.Print "You chose to take a vacation to New York City."
        ElseIf City = 2 Then
            picResults.Print "You chose to take a vacation to Miami."
        ElseIf City = 3 Then
            picResults.Print "You chose to take a vacation to Paris."
        ElseIf City = 4 Then
            picResults.Print "You chose to take a vacation to London."
    End If

'Tells what form of transportation the user chose for their vacation.
    picResults.Print
    If Travel = 1 Then
            picResults.Print "You decided to fly first class to your destination."
        ElseIf Travel = 2 Then
            picResults.Print "You decided to fly business class to your destination."
        ElseIf Travel = 3 Then
            picResults.Print "You decided to fly coach to your destination."
        ElseIf Travel = 4 Then
            picResults.Print "You decided to drive your car to your destination."
        ElseIf Travel = 5 Then
            picResults.Print "You decided to take a bus on your vacation."
    End If
'Tells which place the user wanted to stay
    picResults.Print
    If Room = 1 Then
            picResults.Print "You chose to stay in a condominium."
        ElseIf Room = 2 Then
            picResults.Print "You chose to stay at the luxerious Marriott Hotel."
        ElseIf Room = 3 Then
            picResults.Print "You chose to stay at the Super 8 Hotel."
    End If

'Tells what the user decided to do on their vacation.
    picResults.Print

    Select Case Visit
        Case Is = 1
            picResults.Print "You decided to spend your vacation at the beach.  What a good choice!"
        Case Is = 2
            picResults.Print "You decided to go sight seeing while on vacation."
        Case Else
            picResults.Print "You decided to visit your relatives.  How sweet of you!"
    End Select

    picResults.Print
    picResults.Print
    picResults.Print "Drum rolls please.  Click the TOTAL button to find out how much the vacation will cost."

End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdTotal_Click()
'Gives the user the total cost of the vacation.
    picResults.Print
    picResults.Print "The total cost of your vacation is "; FormatCurrency(Total); "."
    cmdContinue.Visible = True
    cmdTotal.Visible = False
End Sub
