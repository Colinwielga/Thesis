VERSION 5.00
Begin VB.Form frm2003 
   BackColor       =   &H00000080&
   Caption         =   "2003 Champions"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   7680
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   10395
      TabIndex        =   4
      Top             =   5760
      Width           =   10455
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Show Players Name"
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   7680
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   240
      Picture         =   "frm2003.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   10755
      TabIndex        =   1
      Top             =   720
      Width           =   10815
   End
   Begin VB.Label lbl2003Champs 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "2003 NCAA Champions"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frm2003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frm2003
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the team photo from the
'championship season and familiarize the user with names of people on the team.

Option Explicit

Private Sub cmdBack_Click()
    frm2003.Visible = False
    frmChamps.Visible = True
End Sub

Private Sub cmdFill_Click()
Dim Info(1 To 8) As String
Dim Pos As Integer

    picResults.Cls
    
    Open App.Path & "\2003Champs.txt" For Input As #1       'opens text file
    Pos = 0
    
        Do Until Pos = 8            'puts text file into array
            Pos = Pos + 1
            Input #1, Info(Pos)
        Loop
     Close #1
     
    For Pos = 1 To 8                'prints array
        picResults.Print Info(Pos)
    Next Pos
End Sub

Private Sub cmdHome_Click()
    frm2003.Visible = False
    frmMain.Visible = True
End Sub
