VERSION 5.00
Begin VB.Form frm1974 
   BackColor       =   &H00000080&
   Caption         =   "1974 Champions"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFill 
      Caption         =   "Show Players Names"
      Height          =   855
      Index           =   2
      Left            =   5280
      TabIndex        =   5
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   855
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   7200
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   1455
      Left            =   480
      ScaleHeight     =   1395
      ScaleWidth      =   9075
      TabIndex        =   2
      Top             =   5520
      Width           =   9135
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   480
      Picture         =   "frm1974.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   9075
      TabIndex        =   1
      Top             =   840
      Width           =   9135
   End
   Begin VB.Label lbl2002 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "1974 NCAA Champions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   5145
   End
End
Attribute VB_Name = "frm1974"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frm1974
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the team photo from the
'championship season and familiarize the user with names of people on the team.

Option Explicit

Private Sub cmdBack_Click(Index As Integer)
frm1974.Visible = False
frmChamps.Visible = True
End Sub

Private Sub cmdFill_Click(Index As Integer)
Dim Info(1 To 6) As String

Dim Pos As Integer

    picResults.Cls
    
    Open App.Path & "\1974Champs.txt" For Input As #1   'opens the textfile 1974 champs
    Pos = 0
    
        Do Until Pos = 6            'puts the file into array
            Pos = Pos + 1
            Input #1, Info(Pos)
        Loop
     Close #1
     
    For Pos = 1 To 6                'prints the array
        picResults.Print Info(Pos)
    Next Pos
End Sub

Private Sub cmdHome_Click(Index As Integer)
    frm1974.Visible = False     'allows user to access the main screen
    frmMain.Visible = True
End Sub
