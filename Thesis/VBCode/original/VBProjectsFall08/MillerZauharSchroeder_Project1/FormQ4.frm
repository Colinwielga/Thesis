VERSION 5.00
Begin VB.Form FrmQuestion4 
   BackColor       =   &H00800000&
   Caption         =   "Form4"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form4"
   ScaleHeight     =   10575
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   4320
      TabIndex        =   11
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton CmdWalk 
      Caption         =   "Walk Away"
      Height          =   855
      Left            =   1080
      TabIndex        =   10
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton CmdPhone 
      Caption         =   "Phone A Friend"
      Height          =   855
      Left            =   5760
      TabIndex        =   8
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton Cmd50 
      Caption         =   "50/50"
      Height          =   855
      Left            =   2880
      TabIndex        =   7
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton CmdAudience 
      Caption         =   "Ask The Audience"
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "D. Who is Tyler Miller?"
      Height          =   1575
      Left            =   4920
      TabIndex        =   5
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "C. 2"
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton CmdB 
      BackColor       =   &H80000001&
      Caption         =   "B. 1"
      Height          =   1575
      Left            =   4920
      MaskColor       =   &H000000C0&
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton CmdA 
      BackColor       =   &H80000007&
      Caption         =   "A. 0"
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   2895
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
      Left            =   1320
      TabIndex        =   1
      Text            =   "How many dimples are on Tyler Miller's Face?"
      Top             =   2040
      Width           =   5175
   End
   Begin VB.PictureBox PicResults 
      Height          =   1455
      Left            =   3120
      Picture         =   "FormQ4.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FormQ4.frx":0AF8
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
      Left            =   8760
      TabIndex        =   9
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmQuestion4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Who Wants To Be A Millionaire
'Form Name: Question 4
'Authors: Tyler Miller, Troy Zauhar, & Ryan Schroeder
'Date Written: November 2nd, 2008
'Objective: Provide question #4
Option Explicit 'Check for errors

'SEE QUESTION 1 FORM FOR COMPLETE COMMENTORY

Private Sub CmdD_Click()
MsgBox ("Sorry, better luck next time, the correct answer is C.  You won 0 dollars.")

frmQuestion4.Hide
frmEnd.Show
End Sub

Private Sub Cmd50_Click()
CmdA.Enabled = False
CmdB.Enabled = False
FiftyEnabled = False
Cmd50.Enabled = False
End Sub

Private Sub CmdA_Click()
MsgBox ("Sorry, better luck next time, the correct answer is C.  You won 0 dollars.")


frmQuestion4.Hide
frmEnd.Show

End Sub

Private Sub CmdC_Click()
MsgBox ("Correct, let's move on to the $1,000 question")

frmQuestion4.Hide
FrmQuestion5.Show
End Sub

Private Sub CmdB_Click()
MsgBox ("Sorry, better luck next time, the correct answer is C.  You won 0 dollars.")
End
End Sub

Private Sub CmdPhone_Click()
Dim X As String
InputBox ("Who do you want to call? Friend A, B, C, or D?")
If X = "D" Then
MsgBox ("Your friend does not know")
Else
MsgBox ("Your friend thinks C is the right answer, they are 100% sure")
End If
PhoneEnabled = False
CmdPhone.Enabled = False
End Sub
Private Sub CmdAudience_Click()
MsgBox ("Results: A: 15% , B: 40% , C: 35%, D: 10%")
AudienceEnabled = False
CmdAudience.Enabled = False
End Sub
Private Sub cmdquit_Click()
End
End Sub

Private Sub CmdWalk_Click()
MsgBox ("Congratulations on Winning $300!")

frmQuestion4.Hide
frmEnd.Show

End Sub
Private Sub Form_Load()
Cmd50.Enabled = FiftyEnabled
CmdPhone.Enabled = PhoneEnabled
CmdAudience.Enabled = AudienceEnabled
End Sub

