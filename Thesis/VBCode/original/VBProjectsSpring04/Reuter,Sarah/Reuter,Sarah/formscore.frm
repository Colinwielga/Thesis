VERSION 5.00
Begin VB.Form formscore 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form2"
   ClientHeight    =   12675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14940
   LinkTopic       =   "Form2"
   ScaleHeight     =   12675
   ScaleWidth      =   14940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      Caption         =   "Follow me to get the Top Ten list!"
      Enabled         =   0   'False
      Height          =   2295
      Left            =   3360
      Picture         =   "formscore.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Click here to get your score"
      Height          =   2295
      Left            =   600
      Picture         =   "formscore.frx":17F4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFC0&
      Height          =   4455
      Left            =   600
      ScaleHeight     =   4395
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   3840
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   8520
      Width           =   1935
   End
End
Attribute VB_Name = "formscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formscore(formscore.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: score sheet

'switch to next form
Private Sub cmdnext_Click()
formscore.Hide
Formtopten.Show
End Sub

'display how many correct and wrong answers
Private Sub cmdscore_Click()
picresults.Print "You got"; Correct; "of nine questions correct."
picresults.Print
picresults.Print "You got"; Wrong; "of nine questions wrong."
cmdnext.Enabled = True
cmdscore.Enabled = False
End Sub
