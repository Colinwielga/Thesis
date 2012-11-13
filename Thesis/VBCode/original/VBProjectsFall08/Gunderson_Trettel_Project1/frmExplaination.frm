VERSION 5.00
Begin VB.Form frmExplaination 
   BackColor       =   &H000040C0&
   Caption         =   "Cross Country"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   10920
      Picture         =   "frmExplaination.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   10
      Top             =   2160
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   12360
      Picture         =   "frmExplaination.frx":15442
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox txtURL 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "http://www.youtube.com/watch?v=r95NxFw3pyo&feature=related"
      Top             =   8280
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmddirec 
      Caption         =   "Back To Directory"
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play MIAC CC Video"
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.PictureBox picOutput3 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   3360
      ScaleHeight     =   7395
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   480
      Width           =   6855
   End
   Begin VB.CommandButton cmdExplain 
      Caption         =   "What is Cross Country?"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      Caption         =   "November 5, 2008"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   10200
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H000040C0&
      Caption         =   "What is Cross Country"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   10200
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H000040C0&
      Caption         =   "2008 MIAC Cross Country Project "
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   9960
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H000040C0&
      Caption         =   "By: Tyler Trettel and Josh Gunderson"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   9960
      Width           =   2895
   End
End
Attribute VB_Name = "frmExplaination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim SW_SHOW As Boolean, SW_NORMAL As Boolean
'Project Name: MIAC CC Project
'Form Name: frmExplaination
'Authors: Josh Gunderson & Tyler Trettel
'Date: 5 November 2008
'Objective: The purpose of this form is for the user to get a better understanding of how CC operates (Such as the start, finish, course, and scoring)

Private Sub cmddirec_Click()
frmExplaination.Hide
frmdirectory.Show
End Sub

Private Sub cmdExplain_Click()

Dim ccdata(1 To 100) As String, Thing As Integer, T As Integer

Thing = 0
Open App.Path & "\crosscountry.txt" For Input As #1
    
Do Until EOF(1)
    Thing = Thing + 1
    Input #1, ccdata(Thing)
Loop

Close #1

For T = 1 To Thing
    picOutput3.Print ccdata(T)
Next T

End Sub


Private Sub cmdplay_Click()
   Dim URL As String
   
    URL = txtURL.Text
    
    ShellExecute Me.hWnd, "open", URL, "", "", SW_SHOW Or SW_NORMAL
End Sub
