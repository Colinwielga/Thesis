VERSION 5.00
Begin VB.Form frmPoll 
   BackColor       =   &H00808080&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Take A Poll"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10545
   FillColor       =   &H00FF0000&
   FillStyle       =   7  'Diagonal Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   8595
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H8000000D&
      Caption         =   "Go Back"
      Height          =   1695
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdNone 
      BackColor       =   &H000000FF&
      Caption         =   "No Chance At All?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   720
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   7935
   End
   Begin VB.CommandButton cmd50 
      BackColor       =   &H8000000D&
      Caption         =   "50% Chance"
      Height          =   1935
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton cmd75 
      BackColor       =   &H8000000D&
      Caption         =   "75% Chance"
      Height          =   1935
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmd100 
      BackColor       =   &H8000000D&
      Caption         =   "100% Chance"
      Height          =   1935
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmd25 
      BackColor       =   &H8000000D&
      Caption         =   "25% Chance"
      Height          =   1935
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label lblName 
      Caption         =   "By: Chad Henfling"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   8400
      Width           =   1335
   End
End
Attribute VB_Name = "frmPoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Timberwolves Center (MinnesotaTimberwovlesbyChadHenfling.vbp)
'Main Form (frmPoll.frm)
'Chad Henfling
'Created March 23, 2006
'This form allows users to have a little fun in trying to determine the chances the Wolves make the playoffs this year.
'I am having a little fun with the user, basicly telling everyone that the wolves have no chance of winning the playoffs.
Option Explicit
'these commands make all the buttons dissapear in order
Private Sub cmd100_Click()
    cmd100.Visible = False
    cmd75.Visible = True
    cmd50.Visible = False
    cmd25.Visible = False
    cmdNone.Visible = False
    cmdGoBack.Visible = False
End Sub

Private Sub cmd25_Click()
    cmd100.Visible = False
    cmd75.Visible = False
    cmd50.Visible = False
    cmd25.Visible = False
    cmdNone.Visible = True
    cmdGoBack.Visible = False
End Sub

Private Sub cmd50_Click()
    cmd100.Visible = False
    cmd75.Visible = False
    cmd50.Visible = False
    cmd25.Visible = True
    cmdNone.Visible = False
    cmdGoBack.Visible = False
    MsgBox "It is not looking good for the Wolves.......", , "Ooooops!"
End Sub

Private Sub cmd75_Click()
    cmd100.Visible = False
    cmd75.Visible = False
    cmd50.Visible = True
    cmd25.Visible = False
    cmdNone.Visible = False
    cmdGoBack.Visible = False
End Sub

Private Sub cmdGoBack_Click()
    'go back to main form
    frmPoll.Visible = False
    frm1.Visible = True
    cmd100.Visible = True
    cmd75.Visible = False
    cmd50.Visible = False
    cmd25.Visible = False
    cmdNone.Visible = False
    cmdGoBack.Visible = False
End Sub

Private Sub cmdNone_Click()
    Vote = 30
    Vote = Vote + 1
    'displays a message via message box
    MsgBox "You are the " & Vote & " person who believe's the Timberwolves have no chance at making the playoffs", , "Your Results"
    cmdGoBack.Visible = True
    
End Sub


Private Sub Form_Load()
    cmd100.Visible = True
    cmd75.Visible = False
    cmd50.Visible = False
    cmd25.Visible = False
    cmdNone.Visible = False
    cmdGoBack.Visible = False
End Sub


