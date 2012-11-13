VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form5"
   Picture         =   "sources.frx":0000
   ScaleHeight     =   9240
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmainpage 
      Caption         =   "Go back to Main Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   7800
      Width           =   4335
   End
   Begin VB.PictureBox picsources 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      ScaleHeight     =   1875
      ScaleWidth      =   10635
      TabIndex        =   1
      Top             =   5280
      Width           =   10695
   End
   Begin VB.CommandButton cmdsources 
      Caption         =   "Get the Sources!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6960
      TabIndex        =   0
      Top             =   2760
      Width           =   3375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Gilligan's Island
'Form name: sources
'Author:  Emily Olson
'Date written:  March 30, 2008
'Form Objective: display sources used throughout project

Private Sub cmdmainpage_Click()
'load main page
    Form1.Show
    Form5.Hide
End Sub

Private Sub cmdsources_Click()
'display source data from file
    Dim Sources As String
    Open App.Path & "\sources.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Sources
        picsources.Print Sources
    Loop
    Close #1
    
End Sub
