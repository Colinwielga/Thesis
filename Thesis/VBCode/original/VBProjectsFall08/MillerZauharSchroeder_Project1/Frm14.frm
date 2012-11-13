VERSION 5.00
Begin VB.Form FrmQuestion14 
   BackColor       =   &H00800000&
   Caption         =   "Form14"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicResults 
      Height          =   1455
      Left            =   3360
      Picture         =   "Frm14.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   10
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
      TabIndex        =   9
      Text            =   "What continent is subjected to the world's largest ozone hole?"
      Top             =   2160
      Width           =   8535
   End
   Begin VB.CommandButton CmdA 
      BackColor       =   &H80000007&
      Caption         =   "A. North America"
      Height          =   1575
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton CmdB 
      BackColor       =   &H80000001&
      Caption         =   "B. Asia"
      Height          =   1575
      Left            =   5160
      MaskColor       =   &H000000C0&
      TabIndex        =   7
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "C. Antarctica"
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "D.Australia"
      Height          =   1575
      Left            =   5160
      TabIndex        =   5
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdAudience 
      Caption         =   "Ask The Audience"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Cmd50 
      Caption         =   "50/50"
      Height          =   855
      Left            =   3120
      TabIndex        =   3
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton CmdPhone 
      Caption         =   "Phone A Friend"
      Height          =   855
      Left            =   6000
      TabIndex        =   2
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton CmdWalk 
      Caption         =   "Walk Away"
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   9360
      Width           =   2415
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   4560
      TabIndex        =   0
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Frm14.frx":0AF8
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
      TabIndex        =   11
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "FrmQuestion14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Who Wants To Be A Millionaire
'Form Name: Question 14
'Authors: Tyler Miller, Troy Zauhar, & Ryan Schroeder
'Date Written: November 2nd, 2008
'Objective: Provide question #14
Option Explicit

'SEE QUESTION 1 FORM FOR COMPLETE COMMENTORY

Private Sub Cmd50_Click()
CmdA.Enabled = False
CmdD.Enabled = False
FiftyEnabled = False
Cmd50.Enabled = False
End Sub

Private Sub CmdA_Click()
MsgBox ("Sorry, better luck next time, the correct answer was C")
MsgBox ("You are walking away with $25,000! Not too shabby...")

FrmQuestion14.Hide
frmEnd.Show

End Sub

Private Sub CmdAudience_Click()
MsgBox ("Results: A: 15% , B: 50% , C: 2%, D: 33%")
AudienceEnabled = False
CmdAudience.Enabled = False
End Sub

Private Sub CmdB_Click()
MsgBox ("Sorry, better luck next time, the correct answer was C")
MsgBox ("You are walking away with $25,000! Not too shabby...")

FrmQuestion14.Hide
frmEnd.Show

End Sub

Private Sub CmdC_Click()
MsgBox ("Correct, let's move on to the $1,000,000 question!")


FrmQuestion14.Hide
FrmQuestion15.Show
MsgBox ("You have reached the MILLION DOLLAR QUESTION!  Congratulations!  If you answer this question right you will go down in Who Wants To Be A Millionaire HISTORY!  No pressure...")
End Sub

Private Sub CmdD_Click()
MsgBox ("Sorry, better luck next time, the correct answer was C")
MsgBox ("You are walking away with $25,000! Not too shabby...")

FrmQuestion14.Hide
frmEnd.Show

End Sub

Private Sub CmdPhone_Click()
Dim X As String
InputBox ("Who do you want to call? Friend A, B, C, or D?")
If X = "A" Then
MsgBox ("Your friend does not know")
Else
MsgBox ("Your friend thinks B is the right answer, they are 20% sure")
End If
PhoneEnabled = False
CmdPhone.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub CmdWalk_Click()
MsgBox ("Thanks for playing, great job!  You are walking away with $250,000!")

FrmQuestion14.Hide
frmEnd.Show
End Sub

Private Sub Form_Load()
Cmd50.Enabled = FiftyEnabled
CmdAudience.Enabled = AudienceEnabled
CmdPhone.Enabled = PhoneEnabled
End Sub
