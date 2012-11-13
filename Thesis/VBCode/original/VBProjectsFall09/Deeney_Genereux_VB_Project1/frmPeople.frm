VERSION 5.00
Begin VB.Form frmPeople 
   BackColor       =   &H000080FF&
   Caption         =   "They need Help!"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdDiggers 
      Caption         =   "The Diggers!"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdashleys 
      Caption         =   "The Ashley's!"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   1575
      Left            =   3240
      Picture         =   "frmPeople.frx":0000
      Top             =   2040
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   1965
      Left            =   240
      Picture         =   "frmPeople.frx":18CE
      Top             =   1800
      Width           =   2640
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "They can't decide what was better: The Ashley's or The Diggers from ""Recess"".  Which one is better?"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'In this form, the user has to pick between two different
'options to a question, each answer has a different response.

Private Sub cmdashleys_Click()
    MsgBox "All humankind is dissapointed in your decision. You should know better. You didn't help these people. You failed.", , "Bad Choice"
    MsgBox "This is where your story ends. Start over.", , "Story Ends"
    frmPeople.Hide
    frmWelcome.Show
End Sub

Private Sub cmdDiggers_Click()
    MsgBox "Nice Choice! You let those people know who really rocks! The diggers! Here ends your successful journey!", , "Nice Work!"
    MsgBox "This is where your story ends. Start over.", , "Story Ends"
    frmPeople.Hide
    frmWelcome.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
