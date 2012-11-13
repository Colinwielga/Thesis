VERSION 5.00
Begin VB.Form formcover 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   12720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   ScaleHeight     =   12720
   ScaleWidth      =   14895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstart 
      Caption         =   "Click  here to Begin the Quiz"
      Height          =   2535
      Left            =   5880
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sarah's Kermit the Frog Quiz"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Purpose: The purpose of this form is to introduce the user to my program before begining."
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   9000
      Width           =   4455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   $"Form2.frx":1013
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   8160
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Date Written: March 11, 2004"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Form Name: formcover"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Project Name: Kermit's Quiz"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Kermit wishes you the best of luck!"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "You will be able to add your name and score to the top ten list if you do well enough!"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   4080
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "You will be asked nine trivia questions about Kermit the Frog."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Welcome to Sarah's Kermit the Frog quiz for her Computer Science 130 project!"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   3120
      Width           =   5655
   End
End
Attribute VB_Name = "formcover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formcover(Form2.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: cover page


'switch to next form
Private Sub cmdstart_Click()
formcover.Hide
formlive.Show
End Sub
