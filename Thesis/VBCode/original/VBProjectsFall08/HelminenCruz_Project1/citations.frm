VERSION 5.00
Begin VB.Form frmcitations 
   Caption         =   "picresults"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   Picture         =   "citations.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdmainpage 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Go back to Main Page"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Show Citations"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   1560
      ScaleHeight     =   4395
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   2280
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Citations"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmcitations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form inputs the citations and shows them to the user.

Dim citations As String

Private Sub cmdmainpage_Click()
frmcitations.Hide
Welcomeform2.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Command1_Click()

Open App.Path & "\Citations.txt" For Input As #1

    Do Until EOF(1)
        Input #1, citations
        picresults.Print Tab(2); citations
    Loop

Close #1

End Sub

Private Sub Form_Load()

End Sub
