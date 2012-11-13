VERSION 5.00
Begin VB.Form frmBonin 
   BackColor       =   &H00000080&
   Caption         =   "Brian Bonin"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   16
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   15
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Career Total"
      Height          =   735
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Career Average"
      Height          =   735
      Index           =   1
      Left            =   7800
      TabIndex        =   8
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdStats 
      Caption         =   "Show Stats"
      Height          =   735
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   7200
      Width           =   2175
   End
   Begin VB.PictureBox picResults2 
      Height          =   975
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   2955
      TabIndex        =   6
      Top             =   6000
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   4800
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   360
      Picture         =   "frmBonin.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "TP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   7200
      TabIndex        =   14
      Top             =   4440
      Width           =   285
   End
   Begin VB.Label lblAssist 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   6240
      TabIndex        =   13
      Top             =   4440
      Width           =   165
   End
   Begin VB.Label lbGoals 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   5160
      TabIndex        =   12
      Top             =   4440
      Width           =   165
   End
   Begin VB.Label lbGP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "GP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4080
      TabIndex        =   11
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label lblYr 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Yr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label lblBox2 
      BackColor       =   &H00000000&
      Caption         =   $"frmBonin.frx":2607
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3000
      TabIndex        =   4
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Label lblBox1 
      BackColor       =   &H00000000&
      Caption         =   $"frmBonin.frx":282E
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label lblLeo2 
      BackColor       =   &H00000000&
      Caption         =   "Brian Bonin     Center               1996 Award Winner"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblBonin 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Hobey Baker Award Winner Brian Bonin"
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
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   8550
   End
End
Attribute VB_Name = "frmBonin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmBonin
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the user with information about
'Brian Bonin and his accomplishments as a collegiate athlete and professional career.
'The user can view Brian's career statistics, compute totals and averages.

Option Explicit

Dim Year(1 To 4) As String
Dim Games(1 To 4), Goals(1 To 4), Assists(1 To 4), Points(1 To 4) As Integer
Dim Pos As Integer

Private Sub cmdAverage_Click(Index As Integer)
Dim Sum As Integer
Dim Avg As Single

    picResults2.Cls

    For Pos = 1 To 4            'adds up the total games played
        Sum = Sum + Games(Pos)
    Next Pos
    Avg = Sum / 4               'averages the games played
    picResults2.Print "Average Games Played:", Avg      'prints average
    
    Sum = 0
     For Pos = 1 To 4           'adds up the total goals scored
        Sum = Sum + Goals(Pos)
    Next Pos
    Avg = Sum / 4               'averages the goals scored
    picResults2.Print "Average Goals Scored:", Avg      'prints average
    
    Sum = 0
     For Pos = 1 To 4           'adds up the total assists
        Sum = Sum + Assists(Pos)
    Next Pos
    Avg = Sum / 4               'averages the assists
    picResults2.Print "Average Assists:", Avg       'prints average
    
    Sum = 0
     For Pos = 1 To 4           'adds up the total points
        Sum = Sum + Points(Pos)
    Next Pos
    Avg = Sum / 4               'averages the total points
    picResults2.Print "Average Total Points:", Avg      'prints average
End Sub

Private Sub cmdBack_Click(Index As Integer)
    frmBonin.Visible = False
    frmHobeyBaker.Visible = True
End Sub

Private Sub cmdHome_Click(Index As Integer)
    frmBonin.Visible = False
    frmMain.Visible = True
End Sub

Private Sub cmdStats_Click(Index As Integer)
 picResults.Cls
    
    Open App.Path & "\Bonin.txt" For Input As #1        'opens text file
    Pos = 0
    
        Do Until Pos = 4
            Pos = Pos + 1
            Input #1, Year(Pos), Games(Pos), Goals(Pos), Assists(Pos), Points(Pos)  'puts the text file into an array. Each column defined
        Loop
     Close #1
     
    For Pos = 1 To 4
        picResults.Print Year(Pos), Games(Pos), Goals(Pos), Assists(Pos), Points(Pos)   'prints array
    Next Pos
    
    picResults.Print ""
End Sub

Private Sub cmdTotal_Click(Index As Integer)
Dim Sum As Integer

    picResults2.Cls

    For Pos = 1 To 4            'see average comments above
        Sum = Sum + Games(Pos)
    Next Pos
    
    picResults2.Print "Total Games Played:", Sum
    
    Sum = 0
     For Pos = 1 To 4
        Sum = Sum + Goals(Pos)
    Next Pos
    picResults2.Print "Total Goals Scored:", Sum
    
    Sum = 0
     For Pos = 1 To 4
        Sum = Sum + Assists(Pos)
    Next Pos
    picResults2.Print "Total Assists:", , Sum
    
    Sum = 0
     For Pos = 1 To 4
        Sum = Sum + Points(Pos)
    Next Pos
    picResults2.Print "Total Points:", , Sum
    
End Sub
