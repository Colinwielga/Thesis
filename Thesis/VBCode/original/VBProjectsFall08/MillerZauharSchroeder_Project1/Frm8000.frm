VERSION 5.00
Begin VB.Form FrmQuestion8 
   BackColor       =   &H00800000&
   Caption         =   "Form8"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdWalk 
      Caption         =   "Walk Away"
      Height          =   855
      Left            =   1920
      TabIndex        =   11
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   4800
      TabIndex        =   10
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton CmdPhone 
      Caption         =   "Phone A Friend"
      Height          =   855
      Left            =   6360
      TabIndex        =   9
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Cmd50 
      Caption         =   "50/50"
      Height          =   855
      Left            =   3480
      TabIndex        =   8
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton CmdAudience 
      Caption         =   "Ask The Audience"
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "D. George Washington"
      Height          =   1575
      Left            =   5520
      TabIndex        =   6
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "C. JFK"
      Height          =   1575
      Left            =   840
      TabIndex        =   5
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton CmdB 
      BackColor       =   &H80000001&
      Caption         =   "B. Teddy Roosevelt"
      Height          =   1575
      Left            =   5520
      MaskColor       =   &H000000C0&
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton CmdA 
      BackColor       =   &H80000007&
      Caption         =   "A. Thomas Jefferson"
      Height          =   1575
      Left            =   840
      TabIndex        =   3
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
      Left            =   360
      TabIndex        =   2
      Text            =   "Who was the youngest person to ever be the President of the United States?"
      Top             =   2160
      Width           =   8535
   End
   Begin VB.PictureBox PicResults 
      Height          =   1455
      Left            =   3720
      Picture         =   "Frm8000.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Frm8000.frx":0AF8
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
      Left            =   9360
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "FrmQuestion8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Who Wants To Be A Millionaire
'Form Name: Question 8
'Authors: Tyler Miller, Troy Zauhar, & Ryan Schroeder
'Date Written: November 2nd, 2008
'Objective: Provide question #8
Option Explicit 'Check for errors

'SEE QUESTION 1 FORM FOR COMPLETE COMMENTORY

Private Sub CmdD_Click()
MsgBox ("Sorry, better luck next time, the correct answer was B")
MsgBox ("It was not a complete loss for you...you are walking away with $1,000!")

FrmQuestion8.Hide
frmEnd.Show

End Sub

Private Sub Cmd50_Click()
CmdA.Enabled = False
CmdD.Enabled = False
FiftyEnabled = False
Cmd50.Enabled = False


End Sub

Private Sub CmdA_Click()
MsgBox ("Sorry, better luck next time, the correct answer was B")
MsgBox ("It was not a complete loss for you...you are walking away with $1,000!")

FrmQuestion8.Hide
frmEnd.Show

End Sub

Private Sub CmdAudience_Click()
MsgBox ("Results: A: 2% , B: 45% , C: 53%, D: 0%")
AudienceEnabled = False
CmdAudience.Enabled = False
End Sub

Private Sub CmdB_Click()
MsgBox ("Correct, let's move on to the $16,000 question!")


FrmQuestion8.Hide
FrmQuestion9.Show
End Sub

Private Sub CmdC_Click()
MsgBox ("Sorry, better luck next time, the correct answer was B")
MsgBox ("It was not a complete loss for you...you are walking away with $1,000!")

FrmQuestion8.Hide
frmEnd.Show

End Sub

Private Sub CmdPhone_Click()
Dim X As String
InputBox ("Who do you want to call? Friend A, B, C, or D?")
If X = "A" Then
MsgBox ("Your friend thinks the answer is B, they are 85% sure")
Else
MsgBox ("Your friend thinks the answer is B, they are 85% sure")
End If
PhoneEnabled = False
CmdPhone.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub CmdWalk_Click()
MsgBox ("Thanks for playing, great job!  You are walking away with $4,000!")

FrmQuestion8.Hide
frmEnd.Show
End Sub

Private Sub Form_Load()


Cmd50.Enabled = FiftyEnabled
CmdAudience.Enabled = AudienceEnabled
CmdPhone.Enabled = PhoneEnabled

End Sub
