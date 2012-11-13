VERSION 5.00
Begin VB.Form frmSurvey 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H80000015&
      Cancel          =   -1  'True
      Caption         =   "Submit"
      Height          =   800
      Index           =   1
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   2500
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000015&
      Caption         =   "Quit"
      Height          =   800
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   2500
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H80000015&
      Caption         =   "Restart Game"
      Height          =   800
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   2500
   End
   Begin VB.OptionButton Opt5 
      BackColor       =   &H0000FF00&
      Caption         =   "I would trade my first born son for a sequel...."
      Height          =   600
      Left            =   6600
      TabIndex        =   6
      Top             =   6960
      Width           =   2500
   End
   Begin VB.OptionButton Opt4 
      BackColor       =   &H0000FF00&
      Caption         =   "Wow...I mean really...wow..."
      Height          =   600
      Left            =   6600
      TabIndex        =   5
      Top             =   6360
      Width           =   2500
   End
   Begin VB.OptionButton Opt3 
      BackColor       =   &H0000FF00&
      Caption         =   "It makes me more than happy."
      Height          =   600
      Left            =   6600
      TabIndex        =   4
      Top             =   5760
      Width           =   2500
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H0000FF00&
      Caption         =   "There is no other game better than it."
      Height          =   600
      Left            =   6600
      TabIndex        =   3
      Top             =   5160
      Width           =   2500
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H0000FF00&
      Caption         =   "It is the best game ever."
      Height          =   600
      Left            =   6600
      TabIndex        =   0
      Top             =   4560
      Width           =   2500
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Which description below would be describe your experience with Super Awesome Adventure Game?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5040
      TabIndex        =   2
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Survey!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "frmSurvey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Op(1 To 5) As Single
Dim I As Integer


Private Sub cmdQuit_Click(Index As Integer)

    End
    
End Sub

Private Sub cmdStart_Click()

    'Restarts game
    frmSurvey.Visible = False
    frmTitle.Visible = True
    
End Sub

Private Sub cmdSubmit_Click(Index As Integer)

    'Checks user's selections
    
    Open App.Path & "\Survey.txt" For Output As #1
        
            Do While Not EOF(1)
                I = I + 1
                Input #1, Op(I)
            Loop
    
    If opt1.Value = True Then
        Op(1) = Op(1) + 1
    ElseIf Op2.Value = True Then
        Op(2) = Op(2) + 1
    ElseIf Op3.Value = True Then
        Op(3) = Op(3) + 1
    ElseIf Op4.Value = True Then
        Op(4) = Op(4) + 1
    Else
        Op(5) = Op(5) + 1
        
        I = 0
        
            Do While Not EOF(1)
                I = I + 1
                Write #1, Op(I)
            Loop
            
        Close #1
        
        
    
End Sub
