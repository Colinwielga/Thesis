VERSION 5.00
Begin VB.Form FrmQuestion11 
   BackColor       =   &H00800000&
   Caption         =   "Form11"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdWalk 
      Caption         =   "Walk Away"
      Height          =   855
      Left            =   1680
      TabIndex        =   11
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   4560
      TabIndex        =   10
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton CmdPhone 
      Caption         =   "Phone A Friend"
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Cmd50 
      Caption         =   "50/50"
      Height          =   855
      Left            =   3120
      TabIndex        =   7
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton CmdAudience 
      Caption         =   "Ask The Audience"
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "D. Acorn"
      Height          =   1575
      Left            =   5160
      TabIndex        =   5
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "C. Almond"
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdB 
      BackColor       =   &H80000001&
      Caption         =   "B. Peanut"
      Height          =   1575
      Left            =   5160
      MaskColor       =   &H000000C0&
      TabIndex        =   3
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton CmdA 
      BackColor       =   &H80000007&
      Caption         =   "A. Cashew"
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   2895
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
      TabIndex        =   1
      Text            =   "What is the groundnut better known as?"
      Top             =   2160
      Width           =   8535
   End
   Begin VB.PictureBox PicResults 
      Height          =   1455
      Left            =   3360
      Picture         =   "Frm11.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Frm11.frx":0AF8
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
      Width           =   2055
   End
End
Attribute VB_Name = "FrmQuestion11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Who Wants To Be A Millionaire
'Form Name: Question 11
'Authors: Tyler Miller, Troy Zauhar, & Ryan Schroeder
'Date Written: November 2nd, 2008
'Objective: Provide question #11
Option Explicit

'SEE QUESTION 1 FORM FOR COMPLETE COMMENTORY

Private Sub Cmd50_Click()
CmdA.Enabled = False
CmdC.Enabled = False
FiftyEnabled = False
Cmd50.Enabled = False
End Sub

Private Sub CmdA_Click()
MsgBox ("Sorry, better luck next time, the correct answer was B")
MsgBox ("You are walking away with $25,000! Not too shabby...")

FrmQuestion11.Hide
frmEnd.Show

End Sub

Private Sub CmdAudience_Click()
MsgBox ("Results: A: 5% , B: 70% , C: 0%, D: 25%")
AudienceEnabled = False
CmdAudience.Enabled = False
End Sub

Private Sub CmdB_Click()
MsgBox ("Correct, let's move on to the $100,000 question!")

FrmQuestion11.Hide
FrmQuestion12.Show
End Sub

Private Sub CmdC_Click()
MsgBox ("Sorry, better luck next time, the correct answer was B")
MsgBox ("But you are walking away with $25,000...Not too shabby!")

FrmQuestion11.Hide
frmEnd.Show

End Sub

Private Sub CmdD_Click()
MsgBox ("Sorry, better luck next time, the correct answer was B")
MsgBox ("You are walking away with $25,000! Not too shabby...")

FrmQuestion11.Hide
frmEnd.Show

End Sub

Private Sub CmdPhone_Click()
Dim X As String
InputBox ("Who do you want to call? Friend A, B, C, or D?")
If X = "C" Then
MsgBox ("Your friend thinks D is the correct answer, they are 30% sure")
Else
MsgBox ("Your friend thinks B is the right answer, they are 70% sure")
End If
PhoneEnabled = False
CmdPhone.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub CmdWalk_Click()
MsgBox ("That was dumb, it was a free question.  You were walking away with $25,000 right or wrong!")

FrmQuestion11.Hide
frmEnd.Show
End Sub

Private Sub Form_Load()
Cmd50.Enabled = FiftyEnabled
CmdAudience.Enabled = AudienceEnabled
CmdPhone.Enabled = PhoneEnabled
End Sub
