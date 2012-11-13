VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13500
   LinkTopic       =   "Form2"
   Picture         =   "Options.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exit Sacred Heart"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdWorks 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Name a Disease"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdepisode 
      BackColor       =   &H00C0FFC0&
      Caption         =   "View Episodes By Season"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdTrivia 
      BackColor       =   &H00FF80FF&
      Caption         =   "Play Scrubs Trivia"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdMeet 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Meet the Hospital Staff"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Scrubs Project
'Options Menu(frmOptions)
'Ann Boeckmann
'October 25, 2008
'The purpose of this form is to serve as a directory to the activity forms or to leave the hospital
'Each button leads to an activity form or serves as a quit button



Private Sub cmdepisode_Click()

frmOptions.Hide
frmEpisode.Show

End Sub

Private Sub cmdMeet_Click()

frmOptions.Hide
frmStaff.Show

End Sub

Private Sub cmdName_Click()

frmOptions.Hide
frmDisease.Show

End Sub

Private Sub cmdTrivia_Click()

frmOptions.Hide
frmTrivia.Show

End Sub

Private Sub cmdWorks_Click()

frmOptions.Hide
frmWorksCited.Show

End Sub

Private Sub Command1_Click()
MsgBox "Thank you for visiting Sacred Heart!", , "Goodbye!"
End

End Sub
