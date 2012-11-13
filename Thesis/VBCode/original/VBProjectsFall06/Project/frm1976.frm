VERSION 5.00
Begin VB.Form frm1976 
   BackColor       =   &H00000080&
   Caption         =   "1976 Champions"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   855
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Show Players Names"
      Height          =   855
      Index           =   1
      Left            =   5400
      TabIndex        =   4
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   855
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   6720
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   1575
      Left            =   360
      ScaleHeight     =   1515
      ScaleWidth      =   9675
      TabIndex        =   2
      Top             =   4680
      Width           =   9735
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   720
      Picture         =   "frm1976.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   960
      Width           =   9015
   End
   Begin VB.Label lbl1976 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "1976 NCAA Champions"
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
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   5145
   End
End
Attribute VB_Name = "frm1976"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frm1976
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the team photo from the
'championship season and familiarize the user with names of people on the team.

Option Explicit

Private Sub cmdBack_Click(Index As Integer)
frm1976.Visible = False
frmChamps.Visible = True
End Sub

Private Sub cmdFill_Click(Index As Integer)
Dim Info(1 To 7) As String

Dim Pos As Integer


    picResults.Cls
    
    Open App.Path & "\1976Champs.txt" For Input As #1       'opens text file
    Pos = 0
    
        Do Until Pos = 7        'puts text file into array
            Pos = Pos + 1
            Input #1, Info(Pos)
        Loop
     Close #1
     
    For Pos = 1 To 7            'prints the array
        picResults.Print Info(Pos)
    Next Pos
End Sub

Private Sub cmdHome_Click(Index As Integer)
    frm1976.Visible = False
    frmMain.Visible = True
End Sub
