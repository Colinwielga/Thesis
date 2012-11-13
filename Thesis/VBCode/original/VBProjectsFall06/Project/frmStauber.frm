VERSION 5.00
Begin VB.Form frmStauber 
   BackColor       =   &H00000080&
   Caption         =   "Robb Stauber"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   615
      Left            =   360
      TabIndex        =   18
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   360
      TabIndex        =   17
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Career Averages"
      Height          =   735
      Left            =   7680
      TabIndex        =   9
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Career Totals"
      Height          =   735
      Left            =   5520
      TabIndex        =   8
      Top             =   7320
      Width           =   1935
   End
   Begin VB.PictureBox picResults2 
      Height          =   975
      Left            =   3360
      ScaleHeight     =   915
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   6120
      Width           =   3375
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Show Stats"
      Height          =   735
      Left            =   3360
      TabIndex        =   6
      Top             =   7320
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   735
      Left            =   3360
      ScaleHeight     =   675
      ScaleWidth      =   6915
      TabIndex        =   5
      Top             =   5160
      Width           =   6975
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   360
      Picture         =   "frmStauber.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lbSV 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "SV%"
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
      Left            =   9825
      TabIndex        =   16
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label lbGAA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "GAA"
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
      Left            =   8640
      TabIndex        =   15
      Top             =   4800
      Width           =   435
   End
   Begin VB.Label lbGA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "GA"
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
      Left            =   7560
      TabIndex        =   14
      Top             =   4800
      Width           =   315
   End
   Begin VB.Label lbLoss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "L"
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
      Left            =   6600
      TabIndex        =   13
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label lbWins 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "W"
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
      Left            =   5520
      TabIndex        =   12
      Top             =   4800
      Width           =   225
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
      Left            =   4440
      TabIndex        =   11
      Top             =   4800
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
      Left            =   3360
      TabIndex        =   10
      Top             =   4800
      Width           =   315
   End
   Begin VB.Label lblBox2 
      BackColor       =   &H00000000&
      Caption         =   $"frmStauber.frx":8A40
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   6135
   End
   Begin VB.Label lblBox1 
      BackColor       =   &H00000000&
      Caption         =   $"frmStauber.frx":8C85
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   6135
   End
   Begin VB.Label lblLeo2 
      BackColor       =   &H00000000&
      Caption         =   "Robb Stauber Goalie               1988 Award Winner"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblBonin 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Hobey Baker Award Winner Robb Stauber"
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
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmStauber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmStauber
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the user with information about
'Robb Stauber and his accomplishments as a collegiate athlete and professional career.
'The user can view Robb's career statistics, compute totals and averages.


Option Explicit
Dim Year(1 To 3) As String
Dim Wins(1 To 3), Games(1 To 3), Loss(1 To 3), GA(1 To 3), Pos As Integer
Dim GAA(1 To 3), SV(1 To 3) As Single


Private Sub cmdAverage_Click()
Dim Sum As Single
Dim Avg As Single

    picResults2.Cls
    
    Sum = 0                     'to view comments, see frmBonin
     For Pos = 1 To 3
        Sum = Sum + GAA(Pos)
    Next Pos
    Avg = Sum / 3
    picResults2.Print "Career GAA:", FormatNumber(Avg, 2)
    
    Sum = 0
     For Pos = 1 To 3
        Sum = Sum + SV(Pos)
    Next Pos
    Avg = Sum / 3
    picResults2.Print "Career SV%:", FormatNumber(Avg, 3)
    
End Sub

Private Sub cmdBack_Click()
frmStauber.Visible = False
frmHobeyBaker.Visible = True
End Sub

Private Sub cmdFill_Click()
 picResults.Cls
    
    Open App.Path & "\Stauber.txt" For Input As #1
    Pos = 0
    
        Do Until Pos = 3
            Pos = Pos + 1
            Input #1, Year(Pos), Games(Pos), Wins(Pos), Loss(Pos), GA(Pos), GAA(Pos), SV(Pos)
        Loop
     Close #1
     
    For Pos = 1 To 3
        picResults.Print Year(Pos), Games(Pos), Wins(Pos), Loss(Pos), GA(Pos), GAA(Pos), SV(Pos)
    Next Pos
    
    picResults.Print ""
End Sub

Private Sub cmdHome_Click()
frmStauber.Visible = False
frmMain.Visible = True
End Sub

Private Sub cmdTotal_Click()
Dim Sum As Integer

    picResults2.Cls

    For Pos = 1 To 3
        Sum = Sum + Games(Pos)
    Next Pos
    
    picResults2.Print "Total Games Played:", Sum
    
    Sum = 0
     For Pos = 1 To 3
        Sum = Sum + Wins(Pos)
    Next Pos
    picResults2.Print "Total Wins:", , Sum
    
    Sum = 0
     For Pos = 1 To 3
        Sum = Sum + Loss(Pos)
    Next Pos
    picResults2.Print "Total Losses:", , Sum
    
    Sum = 0
     For Pos = 1 To 3
        Sum = Sum + GA(Pos)
    Next Pos
    picResults2.Print "Total Goals Against:", Sum
    
End Sub
