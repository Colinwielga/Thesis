VERSION 5.00
Begin VB.Form frmCSBSJUCrewMain 
   BackColor       =   &H00FFFF00&
   Caption         =   "frmCsb/Sju Crew"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   Picture         =   "frmCSBSJUCrewMain.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWorksCited 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WorksCited"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How Hard are you Working?"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuiz 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Test Your Knowledge"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdMeet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Meet the Members"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdBoat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Basics of Rowing"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CSB/SJU Crew"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   48.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   8175
   End
End
Attribute VB_Name = "frmCSBSJUCrewMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: CSB/SJU Crew
'Form name: frmMeettheMembers
'Authors: Lauren Nephew and Rachel Stalley
'Date: October 8th, 2009
'Objective: Use this form as the home page for the user to exit at or go to the other forms from.
Option Explicit

Private Sub cmdBoat_Click()
'This button allow the user to go to the other forms
frmBoat.Show
frmCSBSJUCrewMain.Hide
End Sub

Private Sub cmdCalculate_Click() 'This button allow the user to go to the other forms
frmCalculateCalories.Show
frmCSBSJUCrewMain.Hide
End Sub

Private Sub cmdMeet_Click() 'This button allow the user to go to the other forms
frmMeettheMembers.Show
frmCSBSJUCrewMain.Hide
End Sub

Private Sub cmdQuit_Click()
'This button allows the user to quit the program
End
End Sub

Private Sub cmdQuiz_Click()
'This button allow the user to go to the other forms
frmQuiz.Show
frmCSBSJUCrewMain.Hide
End Sub

Private Sub cmdWorksCited_Click()
MsgBox "All of our images came off of the CSB/SJU Crew website!", , "Works Cited"
End Sub
