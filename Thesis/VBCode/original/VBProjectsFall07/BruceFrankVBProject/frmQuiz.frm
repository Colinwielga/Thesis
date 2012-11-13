VERSION 5.00
Begin VB.Form frmQuiz 
   Caption         =   "Test Your Timberwolves Knowledge"
   ClientHeight    =   6750
   ClientLeft      =   5670
   ClientTop       =   2505
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   5355
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   8
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "See Results"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   7
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtQ3 
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox txtQ2 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtQ1 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblQ3 
      BackColor       =   &H00FF0000&
      Caption         =   "Question 3: Who's The T-Wolves Coach?"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   4815
   End
   Begin VB.Label lblQ2 
      BackColor       =   &H00FF0000&
      Caption         =   "Queation 2: Who's One Player The T-Wolves Drafted From Villanova?"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblQ1 
      BackColor       =   &H00FF0000&
      Caption         =   "Question 1: What's KG's Nickname?"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblquiz 
      BackColor       =   &H00FF0000&
      Caption         =   "   How Much Do You Know About The Timberwolves? Take The Quiz And Find                           Out!"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   0
      Picture         =   "frmQuiz.frx":0000
      Top             =   0
      Width           =   5370
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdResults_Click()
'This gives the user a Timberwolves quiz
'depending on how many answers the user gets correct a different message box will appear saying something in response
'Case select could have been used in this case but i wanted to use if statements


Dim Dabigticket As String
Dim Florida As String
Dim Coach As String
Dim sum As Integer

Dabigticket = "Da Big Ticket"
Florida = "Randy Foye"
Coach = "Randy Wittman"
sum = 0

If Dabigticket = txtQ1 Then
    sum = sum + 1
End If

If Florida = txtQ2 Then
    sum = sum + 1
End If

If Coach = txtQ3 Then
    sum = sum + 1
End If

If sum = 0 Then
    MsgBox "Zero out of three, you're an idiot"
End If

If sum = 1 Then
    MsgBox "One out three, work on your game"
End If

If sum = 2 Then
    MsgBox "Two out of three, decent but that's not a respectable free throw percentage"
End If

If sum = 3 Then
    MsgBox "Perfect, you're a genius"
End If

    


End Sub

Private Sub cmdreturn_Click()
'This returns the user to the Timberwolves main page
frmQuiz.Visible = False
frmMainPage.Visible = True

End Sub


