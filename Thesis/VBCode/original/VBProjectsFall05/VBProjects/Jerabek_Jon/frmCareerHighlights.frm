VERSION 5.00
Begin VB.Form frmCareerHighlights 
   BackColor       =   &H00800000&
   Caption         =   "Career Highlights"
   ClientHeight    =   6090
   ClientLeft      =   3750
   ClientTop       =   2880
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Impact"
      Size            =   125.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8445
   Visible         =   0   'False
   Begin VB.PictureBox picOutput 
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   125.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   2280
      Picture         =   "frmCareerHighlights.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   8
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton cmdPOW 
      Caption         =   "Player of the Week"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdTripleDouble 
      Caption         =   "Triple- Doubles"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      TabIndex        =   6
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlayerMonth 
      Caption         =   "Player of the Month"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdAllDef 
      Caption         =   "All- Defensive"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdAllNBA 
      Caption         =   "All-NBA"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdMVP 
      Caption         =   "MVP"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdAllStar 
      Caption         =   "All-Star"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdMain3 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   0
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblClick 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Click on any button to reveal the number of times Kevin has received that honor!"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   1920
      TabIndex        =   9
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmCareerHighlights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ProjectKG
'frmCareerHighlights
'Jon Jerabek
'10-26-05
'Objective-Allows user to view # of each stat KG has won

Private Sub cmdAllDef_Click()
picOutput.Cls
picOutput.Print "  6"
End Sub

Private Sub cmdAllNBA_Click()
picOutput.Cls
picOutput.Print "  6"
End Sub

Private Sub cmdAllStar_Click()
picOutput.Cls
picOutput.Print "  8"
End Sub

Private Sub cmdMain3_Click()
frmHome.Show
frmCareerHighlights.Hide
End Sub

Private Sub cmdMVP_Click()
picOutput.Cls
picOutput.Print "  1"
End Sub

Private Sub cmdPlayerMonth_Click()
picOutput.Cls
picOutput.Print "  8"
End Sub

Private Sub cmdPOW_Click()
picOutput.Cls
picOutput.Print " 12"
End Sub

Private Sub cmdTripleDouble_Click()
picOutput.Cls
picOutput.Print " 14"
End Sub
