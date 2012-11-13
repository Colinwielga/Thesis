VERSION 5.00
Begin VB.Form FrmQuestion1 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   FillColor       =   &H00800000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   4920
      TabIndex        =   11
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton CmdWalk 
      Caption         =   "Walk Away"
      Height          =   855
      Left            =   1680
      TabIndex        =   10
      Top             =   9240
      Width           =   2415
   End
   Begin VB.PictureBox PicResults 
      Height          =   1455
      Left            =   3360
      Picture         =   "FormQ1.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Txtorangequestion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Text            =   "What color is an orange?"
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton CmdA 
      BackColor       =   &H80000007&
      Caption         =   "A. Periwinkle"
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton CmdB 
      BackColor       =   &H80000001&
      Caption         =   "B. Orange"
      Height          =   1575
      Left            =   5160
      MaskColor       =   &H000000C0&
      TabIndex        =   5
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "C. Royal Blue"
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "D. Magic Mint"
      Height          =   1575
      Left            =   5160
      TabIndex        =   3
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton CmdAudience 
      Caption         =   "Ask The Audience"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton Cmd50 
      Caption         =   "50/50"
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton CmdPhone 
      Caption         =   "Phone A Friend"
      Height          =   855
      Left            =   6000
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FormQ1.frx":0AF8
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   9975
      Left            =   9000
      TabIndex        =   9
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "FrmQuestion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Who Wants To Be A Millionaire
'Form Name: Question 1
'Authors: Tyler Miller, Troy Zauhar, & Ryan Schroeder
'Date Written: November 2nd, 2008
'Objective: Provide question #1
Option Explicit 'Check for errors

Private Sub CmdD_Click()
MsgBox "Sorry, better luck next time, the correct answer is B.  You won 0 dollars." 'Lets contestant know they answered incorrectly and how much money they won

FrmQuestion1.Hide 'Closes out the current question
frmEnd.Show 'Brings contestant to the end of the game

End Sub

Private Sub Cmd50_Click() '50/50 life-line
CmdA.Enabled = False 'Makes that Letter Button enabled for the remainder of the question
CmdD.Enabled = False 'Makes that Letter Button enabled for the remainder of the question
FiftyEnabled = False 'Makes the 50/50 button disabled for the remainder of the game
Cmd50.Enabled = False 'Makes the 50/50 button disabled for the remainder of the question
End Sub

Private Sub CmdA_Click()
MsgBox "Sorry, better luck next time, the correct answer is B.  You won 0 dollars." 'Lets contestant know they answered incorrectly and how much money they won

FrmQuestion1.Hide 'Closes out the current question
frmEnd.Show 'Brings contestant to the end of the game

End Sub

Private Sub CmdB_Click()
MsgBox "Correct, let's move on to the $200 question" 'Lets the contestant know they answered the question correctly and they are moving on to the next one

FrmQuestion1.Hide 'Closes out the current question
FrmQuestion2.Show 'Brings contestant the the next question
End Sub

Private Sub CmdC_Click()
MsgBox "Sorry, better luck next time, the correct answer is B.  You won 0 dollars." 'Lets contestant know they answered incorrectly and how much money they won

FrmQuestion1.Hide 'Closes out the current question
frmEnd.Show 'Brings contestant to the end of the game

End Sub

Private Sub CmdPhone_Click() ' phone a friend life-line
Dim X As String 'Declares the variable X so it recognizes letters
InputBox ("Who do you want to call? Friend A, B, C, or D?") 'Lets the contestant choose which friend to call
If X = "C" Then 'If the contestant chooses C the following is displayed:
MsgBox ("Your friend does not know")
Else 'If the contestant does not choose C then the following is displayed:
MsgBox ("Your friend thinks B is the right answer, they are 95% sure")
End If 'Ends the If statement
PhoneEnabled = False 'Disables the phone a friend button for the rest of the game
CmdPhone.Enabled = False 'Disbales the phone a friend button for the remainder of the question
End Sub
Private Sub CmdAudience_Click() 'Ask the audience life-line
MsgBox ("Results: A: 15% , B: 50% , C: 35%, D: 0%") 'displays data to the contestant
AudienceEnabled = False 'Disables Audience button for the remainder of the game
CmdAudience.Enabled = False 'Disables Audience button for the raminder of the question
End Sub
Private Sub cmdquit_Click()
End 'Lets the contestant quit the game when they need to
End Sub

Private Sub CmdWalk_Click() 'Walk Away button
MsgBox ("Nice Try") 'Lets the contestant know how much money they are walking away with (if any)

FrmQuestion1.Hide 'Closes out the current question
frmEnd.Show 'Brings contestant to the end of the game
End Sub

Private Sub Form_Load()
Cmd50.Enabled = FiftyEnabled 'Helps with disabling buttons for the entire game after that particular life-line is clicked
CmdPhone.Enabled = PhoneEnabled 'Helps with disabling buttons for the entire game after that particular life-line is clicked
CmdAudience.Enabled = AudienceEnabled 'Helps with disabling buttons for the entire game after that particular life-line is clicked
End Sub

