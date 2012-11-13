VERSION 5.00
Begin VB.Form frmCrew 
   BackColor       =   &H0000C000&
   Caption         =   "Pit Crew = Heart and Soul of the Teams"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form6"
   ScaleHeight     =   8025
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Click to Enter your time"
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   9240
      TabIndex        =   0
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmQuiz.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label lblHeader 
      Caption         =   $"frmQuiz.frx":00AB
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   9255
   End
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   1800
      Picture         =   "frmQuiz.frx":0182
      Top             =   1200
      Width           =   7500
   End
End
Attribute VB_Name = "frmCrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Introduction to NASCAR
'Form Crew
'Colin Roberts and Luke Hommerding
'Written 10/18/09
'The purpose of this form is to give a understanding of a pit crew using a If Then statement
Option Explicit 'declares variables
Dim Time As Single
'This command asks the user to enter a number in order recieve a conditional statement
Private Sub cmdEnter_Click()
    'asks user to enter a time in seconds
    Time = InputBox("Enter your pit crew time in seconds.")
    If Time > 30 Then
        MsgBox ("Invalid Time")
    ElseIf Time >= 25 Then
        MsgBox ("A Pit stop time of ") & Round(Time) & (" Seconds results in: Wow! Big mistake on pit row.  Going to the back of the pack.")
    ElseIf Time >= 20 Then
        MsgBox ("A Pit stop time of ") & Round(Time) & (" Seconds results in: Rough pit stop.  Better make up for it on the next stop.  Lost 15 positions.")
    ElseIf Time >= 15 Then
        MsgBox ("A pit stop time of ") & Round(Time) & (" Second results in: Almost there.  A few mistakes were made.  Lost 4 - 5 positions.")
    ElseIf Time >= 10 Then
        MsgBox ("A pit stop time of ") & Round(Time) & (" Seconds results in: Incredible!  Your pit crew made no mistakes and gained 3 positions.")
    ElseIf Time >= 5 Then
        MsgBox ("A pit stop time of ") & Round(Time) & (" Seconds resutls in: Every pit crew dreams of being this fast, but unfortunately it isn't physically possible.")
    Else
        MsgBox ("Invalid Time")
    End If
       
End Sub
'returns user back to main menu
Private Sub cmdReturn_Click()
    frmMain.Show
    frmCrew.Hide
End Sub
