VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MarioCatcher - Options"
   ClientHeight    =   1692
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1692
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox lstdifficulty 
      Height          =   288
      Left            =   840
      TabIndex        =   4
      Text            =   "Select a level of difficulty"
      Top             =   240
      Width           =   3252
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1932
   End
   Begin VB.Label lblDifficulty 
      Caption         =   "Difficulty:"
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   732
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmoptions
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to select the level they want to play the mario catcher game.  They can
                'select from seven difficulties  or return to the main page.  The purpose of the form to to prepare for the game

Option Explicit

Private Sub cmdCancel_Click()
    frmOptions.Hide     'hides the options page
    frmMariocatcher.Hide        'hides the mario catcher page
    frmMain.Show        'shows the main page
End Sub

Private Sub cmdOK_Click()
    Select Case lstdifficulty.ListIndex     'looks at the difficulty selected and finds its index
        Case Is = 0     'if you select the easiest, the lowest index, the timer is set according (slowest timer)
            frmMariocatcher.GlobTimeout = 10        'the global timer is set
            frmMariocatcher.MarioTimeout = 100      'the timer for the moving mario is set
        Case Is = 1     'checks the index and loops if it matches the case
            frmMariocatcher.GlobTimeout = 10        'sets the global timer
            frmMariocatcher.MarioTimeout = 90       'sets timer for the moving mario
        Case Is = 2     'checks the index and loops if it matches the case
            frmMariocatcher.GlobTimeout = 10        'sets the global timer
            frmMariocatcher.MarioTimeout = 70       'sets timer for the moving mario
        Case Is = 3     'checks the index and loops if it matches the case
            frmMariocatcher.GlobTimeout = 10        'sets the global timer
            frmMariocatcher.MarioTimeout = 60       'sets timer for the moving mario
        Case Is = 4     'checks the index and loops if it matches the case
            frmMariocatcher.GlobTimeout = 10        'sets the global timer
            frmMariocatcher.MarioTimeout = 40       'sets timer for the moving mario
        Case Is = 5     'checks the index and loops if it matches the case
            frmMariocatcher.GlobTimeout = 10        'sets the global timer
            frmMariocatcher.MarioTimeout = 20       'sets timer for the moving mario
        Case Is = 6     'checks the index and loops if it matches the case
            frmMariocatcher.GlobTimeout = 5     'sets the global timer
            frmMariocatcher.MarioTimeout = 10       'sets timer for the moving mario
        Case Is = 7     'checks the index and loops if it matches the case
            frmMariocatcher.GlobTimeout = 1     'sets the global timer
            frmMariocatcher.MarioTimeout = 5        'sets timer for the moving mario
    End Select
    frmOptions.Hide     'hides the options page
    frmMariocatcher.Show        'shows the mario catcher page
End Sub

Private Sub Form_Load()
    
    With lstdifficulty      'when the form is loaded, these varying difficulties are displayed in the picture box for the user
        .AddItem "Totally Easy", 0      'index's the difficulties in order of easiest to hardest
        .AddItem "Very Easy", 1
        .AddItem "Somewhat Easy", 2
        .AddItem "Easy", 3
        .AddItem "Moderate", 4
        .AddItem "Experienced Catcher", 5
        .AddItem "Master of the Marios", 6
        .AddItem "Can you get any better?", 7
    End With
    
End Sub

