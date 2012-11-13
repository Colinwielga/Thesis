VERSION 5.00
Begin VB.Form frmStartup 
   Caption         =   "Start Up"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   Picture         =   "frmStartup.frx":0000
   ScaleHeight     =   10785
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9240
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00008000&
      Caption         =   "Start Program"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9240
      Width           =   1575
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MN Deer'
'Form Name: Startup'
'Authors: Jordon Przybilla'
'Date Written: October 4, 2009
'this program will provide basic minnesota whitetail deer information but will focus mainly on deer hunting'
'deer specific terms will be defined and one form will allow the user to pick items they need to start'
'deer hunting and will give them a ruff estimate of how much they will need to spend to get into hunting'

'this form is basically just to initialize the program'

Option Explicit



Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStart_Click()

'this button will initialize the program. before moving on to the other forms though there will be a trivia question'
'after the user has entered a guess and the msgbox has popped up the program will move to the next form.

Dim pop As Long
'trivia question
pop = InputBox("Let's start with a trivia question.  Take a guess at the deer population in MN every spring after the fawns are born.", , 0)
If pop > 900000 And pop < 1000000 Then
    MsgBox "Dang you're smart, according to the Minnesota DNR the deer population every spring is between 900,000 and 1,000,000.", , "Correct"
ElseIf pop > 750000 And pop < 1250000 Then
    MsgBox "Close, but you're still wrong, according to the Minnesota DNR the deer population every spring is between 900,000 and 1,000,000.", , "Almost"
Else
    MsgBox "Sorry but you aren't even close.  According to the Minnesota DNR the deer population every spring reaches between 900,000 and 1,000,000.", , "You're wrong"
End If

frmStartup.Hide
frmHome.Show

End Sub

