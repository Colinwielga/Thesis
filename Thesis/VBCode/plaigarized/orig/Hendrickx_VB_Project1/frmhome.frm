VERSION 5.00
Begin VB.Form frmhome 
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   4845
   ClientTop       =   2595
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   Picture         =   "frmhome.frx":0000
   ScaleHeight     =   6012.396
   ScaleMode       =   0  'User
   ScaleWidth      =   11835
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00FF8080&
      Caption         =   "       Drink Me?  Eat Me?  No, Click Me!"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdcited 
      Caption         =   "Cited Sources"
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdquiz 
      BackColor       =   &H00FF8080&
      Caption         =   "How well do you know your Disney characters?"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdtrivia 
      BackColor       =   &H00FF8080&
      Caption         =   "Fun Facts"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdlist 
      BackColor       =   &H00FF8080&
      Caption         =   "What animated movies did Disney create?"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      Caption         =   "The Wonderful World of Disney"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   44.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2655
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   11175
   End
End
Attribute VB_Name = "frmhome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Wonderful World of Disney
'form Home
'Kate Hendrickx
'February 2010
'Objective: this form is "home base"--in order to move on to another form,
'the user has to go back to this form first.
'Goal: to have a fun program for anyone of any age to use and enjoy.
Option Explicit

Private Sub cmdlist_Click()
frmarray.Show
frmhome.Hide
End Sub

Private Sub cmdcited_Click()
frmcited.Show
frmhome.Hide
End Sub

Private Sub cmdquiz_Click()
frmquiz.Show
frmhome.Hide
End Sub

Private Sub cmdstart_Click()
Dim yourname As String

yourname = InputBox("What is your name?")
    If yourname = " " Then
    MsgBox "Error: please use letters only."
Else: MsgBox "Welcome " & yourname & "!  To get to Neverland, fly to the second star on the right, straight on till morning!"
End If

cmdquiz.Visible = True
cmdlist.Visible = True
cmdcited.Visible = True
cmdquit.Visible = True
cmdtrivia.Visible = True

End Sub

Private Sub cmdtrivia_Click()
frmtrivia.Show
frmhome.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub
