VERSION 5.00
Begin VB.Form frmDancer 
   Caption         =   "Win a Date With a Dancer!"
   ClientHeight    =   5910
   ClientLeft      =   3585
   ClientTop       =   2715
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8610
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmdDancer1 
      Caption         =   "See If You Can Win a Date With a Timberwolves Dancer!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   5910
      Left            =   0
      Picture         =   "frmDancer.frx":0000
      Top             =   0
      Width           =   8610
   End
End
Attribute VB_Name = "frmDancer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDancer1_Click()
'The user is asked to pick a number between one and one hundred based on their answer they could win a date with one of four timberwolve dancers or a date with crunch
'The user inputs their guess through an input box
'Case select is used because their are five possible outcomes
'After making their guess the user is taken to a different form depending on their guess which tells them who they get to go out with

Dim Number As Single
Number = InputBox("Pick a number between 1 and 100")
Select Case Number
    Case 1 To 20
        frmDancer.Visible = False
        frmBianca.Visible = True
        MsgBox "Nice choice, your going out with Bianca!"
    Case 20 To 40
        frmDancer.Visible = False
        frmBianca.Visible = True
        MsgBox "Ooh La La, You get to go out with Kristi"
    Case 40 To 60
        frmDancer.Visible = False
        frmBianca.Visible = True
        MsgBox "Time to get your swerve on with Shalisha"
    Case 60 To 80
        frmDancer.Visible = False
        frmBianca.Visible = True
        MsgBox "Nice one, you and Kristi have all night together."
    Case 80 To 100
        frmDancer.Visible = False
        frmCrunch.Visible = True
        MsgBox "You Fool! You're spending the night with Crunch."
    Case Else
        MsgBox "Enter a number between 1 and 100"
    End Select
    


End Sub

Private Sub cmdreturn_Click()
'This returns the user to the main page form and away from the dancer form

frmDancer.Visible = False
frmMainPage.Visible = True

End Sub
