VERSION 5.00
Begin VB.Form frm2002 
   BackColor       =   &H00000080&
   Caption         =   "2002 Champions"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFill 
      Caption         =   "Show Players Names"
      Height          =   735
      Left            =   5280
      TabIndex        =   5
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   6960
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   1695
      Left            =   360
      ScaleHeight     =   1635
      ScaleWidth      =   9675
      TabIndex        =   2
      Top             =   5160
      Width           =   9735
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   720
      Picture         =   "frm2002.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   9075
      TabIndex        =   1
      Top             =   840
      Width           =   9135
   End
   Begin VB.Label lbl2002 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "2002 NCAA Champions"
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
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frm2002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frm2002
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the team photo from the
'championship season and familiarize the user with names of people on the team.

Option Explicit

Private Sub cmdBack_Click()
frm2002.Visible = False
frmChamps.Visible = True
End Sub

Private Sub cmdFill_Click()
Dim Info(1 To 8) As String
Dim Pos As Integer

    picResults.Cls
    
    Open App.Path & "\2002Champs.txt" For Input As #1       'opens the text file
    Pos = 0
    
        Do Until Pos = 8            'puts text file into an array
            Pos = Pos + 1
            Input #1, Info(Pos)
        Loop
     Close #1
     
    For Pos = 1 To 8                'prints array
        picResults.Print Info(Pos)
    Next Pos
End Sub

Private Sub cmdHome_Click()
frm2002.Visible = False
frmMain.Visible = True
End Sub
