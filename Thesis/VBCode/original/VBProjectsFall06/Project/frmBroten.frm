VERSION 5.00
Begin VB.Form frmBroten 
   BackColor       =   &H00000080&
   Caption         =   "Neal Broten"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   6960
      Width           =   1575
   End
   Begin VB.PictureBox picResults2 
      Height          =   855
      Left            =   4200
      ScaleHeight     =   795
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   6480
      Width           =   3255
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Career Averages"
      Height          =   735
      Left            =   8280
      TabIndex        =   8
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Career Totals"
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Show Stats"
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   7560
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   615
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   5640
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   240
      Picture         =   "frmBroten.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   960
      Width           =   3735
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
      Left            =   8400
      TabIndex        =   14
      Top             =   5280
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
      Left            =   7440
      TabIndex        =   13
      Top             =   5280
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
      Left            =   6360
      TabIndex        =   12
      Top             =   5280
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
      Left            =   5280
      TabIndex        =   11
      Top             =   5280
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
      Left            =   4200
      TabIndex        =   10
      Top             =   5280
      Width           =   315
   End
   Begin VB.Label lblBox2 
      BackColor       =   &H00000000&
      Caption         =   $"frmBroten.frx":7A90
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   4200
      TabIndex        =   4
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Label lblBox1 
      BackColor       =   &H00000000&
      Caption         =   $"frmBroten.frx":7E0A
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   4200
      TabIndex        =   3
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label lblBroten2 
      BackColor       =   &H00000000&
      Caption         =   "Neal Broten     Center               1981 Award Winner"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblBroten 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Hobey Baker Award Winner Neal Broten"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   8595
   End
End
Attribute VB_Name = "frmBroten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmBroten
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the user with information about
'Neal Broten and his accomplishments as a collegiate athlete and professional career.
'The user can view Neal's career statistics, compute totals and averages.


Option Explicit
Dim Year(1 To 2) As String
Dim Games(1 To 2), Goals(1 To 2), Assists(1 To 2), Points(1 To 2), Pos As Integer

Private Sub cmdAverage_Click()
Dim Sum As Integer
Dim Avg As Single

    picResults2.Cls

    For Pos = 1 To 2            'to view comments, see frmBonin
        Sum = Sum + Games(Pos)
    Next Pos
    Avg = Sum / 2
    picResults2.Print "Average Games Played:", Avg
    
    Sum = 0
     For Pos = 1 To 2
        Sum = Sum + Goals(Pos)
    Next Pos
    Avg = Sum / 2
    picResults2.Print "Average Goals Scored:", Avg
    
    Sum = 0
     For Pos = 1 To 2
        Sum = Sum + Assists(Pos)
    Next Pos
    Avg = Sum / 2
    picResults2.Print "Average Assists:", Avg
    
    Sum = 0
     For Pos = 1 To 2
        Sum = Sum + Points(Pos)
    Next Pos
    Avg = Sum / 2
    picResults2.Print "Average Total Points:", Avg
End Sub

Private Sub cmdBack_Click()
frmBroten.Visible = False
frmHobeyBaker.Visible = True
End Sub

Private Sub cmdFill_Click()

    picResults.Cls
    
    Open App.Path & "\Broten.txt" For Input As #1
    Pos = 0
    
        Do Until Pos = 2
            Pos = Pos + 1
            Input #1, Year(Pos), Games(Pos), Goals(Pos), Assists(Pos), Points(Pos)
        Loop
     Close #1
     
    For Pos = 1 To 2
        picResults.Print Year(Pos), Games(Pos), Goals(Pos), Assists(Pos), Points(Pos)
    Next Pos
    
    picResults.Print ""
End Sub

Private Sub cmdHome_Click()
frmBroten.Visible = False
frmMain.Visible = True
End Sub

Private Sub cmdTotal_Click()
Dim Sum As Integer

    picResults2.Cls

    For Pos = 1 To 2
        Sum = Sum + Games(Pos)
    Next Pos
    
    picResults2.Print "Total Games Played:", Sum
    
    Sum = 0
     For Pos = 1 To 2
        Sum = Sum + Goals(Pos)
    Next Pos
    picResults2.Print "Total Goals Scored:", Sum
    
    Sum = 0
     For Pos = 1 To 2
        Sum = Sum + Assists(Pos)
    Next Pos
    picResults2.Print "Total Assists:", , Sum
    
    Sum = 0
     For Pos = 1 To 2
        Sum = Sum + Points(Pos)
    Next Pos
    picResults2.Print "Total Points:", , Sum
End Sub
