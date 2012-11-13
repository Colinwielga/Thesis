VERSION 5.00
Begin VB.Form frm1979 
   BackColor       =   &H00000080&
   Caption         =   "1979 Champions"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   855
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Show Players Names"
      Height          =   855
      Index           =   1
      Left            =   5520
      TabIndex        =   4
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   855
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   6600
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   1455
      Left            =   480
      ScaleHeight     =   1395
      ScaleWidth      =   9435
      TabIndex        =   2
      Top             =   4800
      Width           =   9495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   3375
      Left            =   840
      Picture         =   "frm1979.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   8715
      TabIndex        =   1
      Top             =   1200
      Width           =   8775
   End
   Begin VB.Label lbl1976 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "1979 NCAA Champions"
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
      Top             =   360
      Width           =   5145
   End
End
Attribute VB_Name = "frm1979"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frm1979
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the team photo from
'the championship season and familiarize the user with names of people on the team.

Private Sub cmdBack_Click(Index As Integer)
frm1979.Visible = False
frmChamps.Visible = True
End Sub

Private Sub cmdFill_Click(Index As Integer)
Dim Info(1 To 7) As String

Dim Pos As Integer


    picResults.Cls
    
    Open App.Path & "\1979Champs.txt" For Input As #1       'opens the text file
    Pos = 0
    
        Do Until Pos = 7            'puts text file into an array
            Pos = Pos + 1
            Input #1, Info(Pos)
        Loop
     Close #1
        
    For Pos = 1 To 7                'prints the array
        picResults.Print Info(Pos)
    Next Pos
End Sub

Private Sub cmdHome_Click(Index As Integer)
frm1979.Visible = False
frmMain.Visible = True
End Sub
