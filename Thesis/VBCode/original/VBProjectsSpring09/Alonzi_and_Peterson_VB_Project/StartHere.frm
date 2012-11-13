VERSION 5.00
Begin VB.Form StartHere 
   BackColor       =   &H00FFFF80&
   Caption         =   "Ready to Go?"
   ClientHeight    =   2205
   ClientLeft      =   6420
   ClientTop       =   4635
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4875
   Begin VB.CommandButton Nope 
      BackColor       =   &H008080FF&
      Caption         =   "Not yet."
      Height          =   855
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Yep 
      BackColor       =   &H008080FF&
      Caption         =   "Yes!"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Have you used the VB piano before?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "StartHere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Palonzison Piano
'This is the StartHere Form
'Matthew Peterson and Nicholas Alonzi are the authors of this Form
'This form was written in 2009 in the month of March
'This form is to give the option to go to the help page for first time users
    'and if you don't need help, it will skip the page.
'This page compiles the array for the piano note files or for the instructions
    'depending on what page you navigate to next
'The purpose of this project is to provide a visually pleasing program
    'that lets the user play with the piano in a semirealistic way
    'and provides a couple of songs for those with more interest in music
    
Private Sub Nope_Click()
    Open App.Path & ("\instructions.txt") For Input As #1
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, InsNum(Ctr), Instruction(Ctr)
    Loop
    Close #1
    StartHere.Hide
    Instructions.Show
End Sub

Private Sub Yep_Click()
Dim TempNote(1 To 99) As String
    Open App.Path & "\notefiles.txt" For Input As #1
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, TempNote(Ctr)
        Notes(Ctr) = App.Path & TempNote(Ctr)
    Loop
    Close #1
    StartHere.Hide
    Piano.Show
End Sub

