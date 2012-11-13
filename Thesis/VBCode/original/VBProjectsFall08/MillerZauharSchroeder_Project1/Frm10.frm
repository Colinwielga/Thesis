VERSION 5.00
Begin VB.Form FrmQuestion10 
   BackColor       =   &H00800000&
   Caption         =   "Form10"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdWalk 
      Caption         =   "Walk Away"
      Height          =   855
      Left            =   1680
      TabIndex        =   11
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   4560
      TabIndex        =   10
      Top             =   9000
      Width           =   2415
   End
   Begin VB.PictureBox PicResults 
      Height          =   1455
      Left            =   3360
      Picture         =   "Frm10.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Txt1 
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
      Left            =   0
      TabIndex        =   7
      Text            =   "What is the symbol for Iron?"
      Top             =   2160
      Width           =   8535
   End
   Begin VB.CommandButton CmdA 
      BackColor       =   &H80000007&
      Caption         =   "A. Fe"
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton CmdB 
      BackColor       =   &H80000001&
      Caption         =   "B. Ir"
      Height          =   1575
      Left            =   5160
      MaskColor       =   &H000000C0&
      TabIndex        =   5
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "C. Au"
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "D. Ps"
      Height          =   1575
      Left            =   5160
      TabIndex        =   3
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdAudience 
      Caption         =   "Ask The Audience"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Cmd50 
      Caption         =   "50/50"
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton CmdPhone 
      Caption         =   "Phone A Friend"
      Height          =   855
      Left            =   6000
      TabIndex        =   0
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Frm10.frx":0AF8
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
      Height          =   10095
      Left            =   9000
      TabIndex        =   9
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "FrmQuestion10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Who Wants To Be A Millionaire
'Form Name: Question 10
'Authors: Tyler Miller, Troy Zauhar, & Ryan Schroeder
'Date Written: November 2nd, 2008
'Objective: Provide question #10
Option Explicit

'SEE QUESTION 1 FORM FOR COMPLETE COMMENTORY

Private Sub Cmd50_Click()
CmdD.Enabled = False
CmdC.Enabled = False
FiftyEnabled = False
Cmd50.Enabled = False
End Sub

Private Sub CmdA_Click()
MsgBox ("Correct, you now have $25,000 no matter what!")


MsgBox ("Let's move on to the $50,000 question!")
FrmQuestion10.Hide
FrmQuestion11.Show
End Sub

Private Sub CmdAudience_Click()
MsgBox ("Results: A: 30% , B: 42% , C: 8%, D: 20%")
AudienceEnabled = False
CmdAudience.Enabled = False
End Sub

Private Sub CmdB_Click()
MsgBox ("Sorry, better luck next time, the correct answer was A")
MsgBox ("It was not a complete loss for you...you are walking away with $1,000!")

FrmQuestion10.Hide
frmEnd.Show

End Sub

Private Sub CmdC_Click()
MsgBox ("Sorry, better luck next time, the correct answer was A")
MsgBox ("It was not a complete loss for you...you are walking away with $1,000!")

FrmQuestion10.Hide
frmEnd.Show

End Sub

Private Sub CmdD_Click()
MsgBox ("Sorry, better luck next time, the correct answer was A")
MsgBox ("It was not a complete loss for you...you are walking away with $1,000!")

FrmQuestion10.Hide
frmEnd.Show
End Sub

Private Sub CmdPhone_Click()
Dim X As String
InputBox ("Who do you want to call? Friend A, B, C, or D?")
If X = "C" Then
MsgBox ("Your friend thinks B is the correct answer, they are 65% sure")
Else
MsgBox ("Your friend thinks A is the right answer, they are 80% sure")
End If
PhoneEnabled = False
CmdPhone.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub CmdWalk_Click()
MsgBox ("Thanks for playing, great job!  You are walking away with $16,000!")

FrmQuestion10.Hide
frmEnd.Show

End Sub

Private Sub Form_Load()
Cmd50.Enabled = FiftyEnabled
CmdAudience.Enabled = AudienceEnabled
CmdPhone.Enabled = PhoneEnabled

End Sub
