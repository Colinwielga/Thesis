VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome!"
   ClientHeight    =   8880
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   Picture         =   "frmWelcome.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   14955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   3
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmdbegin 
      BackColor       =   &H00008000&
      Caption         =   "Start The Adventure!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12120
      MaskColor       =   &H00008000&
      Picture         =   "frmWelcome.frx":249F42
      TabIndex        =   2
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Where the Story Becomes Your Own"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1335
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   9615
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Mystic Forest!"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Katie Deeney & Elise Generex
'Create Your own ending
'Date Done: 10/10/2009
'Welcome
'This is the welcoming page to our project
'Our project: "Create your own ending" Takes you through the journeys of 4 characters
'You are able to choose your own path on some journeys
Private Sub cmdbegin_Click()
    Open App.Path & "\Bios.txt" For Input As #1 'Opens the file
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Names(Ctr), Strengths(Ctr), Weaknesses(Ctr)
    Loop
    Close #1
    frmWelcome.Hide 'Goes to the next form
    frmCharacters.Show
    
End Sub

Private Sub CmdQuit_Click()
    End 'Ends the program
End Sub
