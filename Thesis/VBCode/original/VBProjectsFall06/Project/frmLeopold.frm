VERSION 5.00
Begin VB.Form frmLeopold 
   BackColor       =   &H00000080&
   Caption         =   "Jordan Leopold"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   615
      Left            =   600
      TabIndex        =   16
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   600
      TabIndex        =   15
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox picResults2 
      Height          =   975
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   2955
      TabIndex        =   9
      Top             =   6240
      Width           =   3015
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Career Average"
      Height          =   615
      Left            =   7320
      TabIndex        =   8
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Totals"
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Show Stats"
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   7440
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   5040
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   480
      Picture         =   "frmLeopold.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
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
      Top             =   4680
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
      Top             =   4680
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
      Top             =   4680
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
      Top             =   4680
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
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label lblBox2 
      BackColor       =   &H00000000&
      Caption         =   $"frmLeopold.frx":2AF7
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   6975
   End
   Begin VB.Label lblBox1 
      BackColor       =   &H00000000&
      Caption         =   $"frmLeopold.frx":2D33
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label lblLeo2 
      BackColor       =   &H00000000&
      Caption         =   "Jordan Leopold Defenseman        2002 Award Winner"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblLeo 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Hobey Baker Award Winner Jordan Leopold"
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
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   9390
   End
End
Attribute VB_Name = "frmLeopold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmLeopold
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the user with information about
'Jordan Leopold and his accomplishments as a collegiate athlete and professional career.
'The user can view Jordan's career statistics, compute totals and averages.


Option Explicit
Dim Year(1 To 4) As String
Dim Games(1 To 4), Goals(1 To 4), Assists(1 To 4), Points(1 To 4), Pos As Integer


Private Sub cmdAverage_Click()

Dim Sum As Integer
Dim Avg As Single

    picResults2.Cls

    For Pos = 1 To 4            'to view comments, see frmBonin
        Sum = Sum + Games(Pos)
    Next Pos
    Avg = Sum / 4
    picResults2.Print "Average Games Played:", Avg
    
    Sum = 0
     For Pos = 1 To 4
        Sum = Sum + Goals(Pos)
    Next Pos
    Avg = Sum / 4
    picResults2.Print "Average Goals Scored:", Avg
    
    Sum = 0
     For Pos = 1 To 4
        Sum = Sum + Assists(Pos)
    Next Pos
    Avg = Sum / 4
    picResults2.Print "Average Assists:", Avg
    
    Sum = 0
     For Pos = 1 To 4
        Sum = Sum + Points(Pos)
    Next Pos
    Avg = Sum / 4
    picResults2.Print "Average Total Points:", Avg
End Sub

Private Sub cmdBack_Click()
frmLeopold.Visible = False
frmHobeyBaker.Visible = True
End Sub

Private Sub cmdFill_Click()

    picResults.Cls
    
    Open App.Path & "\Leopold.txt" For Input As #1
    Pos = 0
    
        Do Until Pos = 4
            Pos = Pos + 1
            Input #1, Year(Pos), Games(Pos), Goals(Pos), Assists(Pos), Points(Pos)
        Loop
     Close #1
     
    For Pos = 1 To 4
        picResults.Print Year(Pos), Games(Pos), Goals(Pos), Assists(Pos), Points(Pos)
    Next Pos
    
    picResults.Print ""
End Sub

Private Sub cmdHome_Click()
frmLeopold.Visible = False
frmMain.Visible = True
End Sub

Private Sub cmdTotal_Click()
Dim Sum As Integer

    picResults2.Cls

    For Pos = 1 To 4
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
