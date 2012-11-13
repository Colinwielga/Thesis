VERSION 5.00
Begin VB.Form FrmQuestion5 
   BackColor       =   &H00800000&
   Caption         =   "Form5"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13950
   LinkTopic       =   "Form5"
   ScaleHeight     =   10770
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   4440
      TabIndex        =   11
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton CmdWalk 
      Caption         =   "Walk Away"
      Height          =   855
      Left            =   1200
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
      Caption         =   "D. The University of Minnesota Marching Band"
      Height          =   1575
      Left            =   4920
      TabIndex        =   5
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "C. Little Big Town"
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton CmdB 
      BackColor       =   &H80000001&
      Caption         =   "B. Nappy Roots"
      Height          =   1575
      Left            =   4920
      MaskColor       =   &H000000C0&
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton CmdA 
      BackColor       =   &H80000007&
      Caption         =   "A. ACDC"
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
      Text            =   "Which one of these bands is from Minnesota?"
      Top             =   2040
      Width           =   4935
   End
   Begin VB.PictureBox PicResults 
      Height          =   1455
      Left            =   3120
      Picture         =   "FormQ5.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FormQ5.frx":0AF8
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
Attribute VB_Name = "FrmQuestion5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Who Wants To Be A Millionaire
'Form Name: Question 5
'Authors: Tyler Miller, Troy Zauhar, & Ryan Schroeder
'Date Written: November 2nd, 2008
'Objective: Provide question #5
Option Explicit 'Check for errors

'SEE QUESTION 1 FORM FOR COMPLETE COMMENTORY

Private Sub CmdC_Click()
MsgBox ("Sorry, better luck next time, the correct answer is D.  You won 0 dollars.")

FrmQuestion5.Hide
frmEnd.Show

End Sub

Private Sub Cmd50_Click()
CmdA.Enabled = False
CmdC.Enabled = False
FiftyEnabled = False
Cmd50.Enabled = False
End Sub

Private Sub CmdA_Click()
MsgBox ("Sorry, better luck next time, the correct answer is D.  You won 0 dollars.")


FrmQuestion5.Hide
frmEnd.Show
End Sub

Private Sub CmdD_Click()
MsgBox ("Correct, you have won $1,000!")

FrmQuestion5.Hide
FrmQuestion6.Show
End Sub

Private Sub CmdB_Click()
MsgBox ("Sorry, better luck next time, the correct answer is D.  You won 0 dollars.")

FrmQuestion5.Hide
frmEnd.Show
End Sub

Private Sub CmdPhone_Click()
Dim X As String
InputBox ("Who do you want to call? Friend A, B, C, or D?")
If X = "A" Then
MsgBox ("Your friend does not know")
Else
MsgBox ("Your friend thinks D is the right answer, they are 70% sure")
End If
PhoneEnabled = False
CmdPhone.Enabled = False
End Sub
Private Sub CmdAudience_Click()
MsgBox ("Results: A: 25% , B: 15% , C: 35%, D: 25%")
AudienceEnabled = False
CmdAudience.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub CmdWalk_Click()
MsgBox ("Congratulations on Winning $500")

FrmQuestion5.Hide
frmEnd.Show

End Sub
Private Sub Form_Load()
Cmd50.Enabled = FiftyEnabled
CmdPhone.Enabled = PhoneEnabled
CmdAudience.Enabled = AudienceEnabled
MsgBox ("Remember, if you answer this question right you will walk away with no less than $1,000!")
End Sub

