VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FF8080&
   Caption         =   "Form7"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form7"
   ScaleHeight     =   7515
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H00FF8080&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      TabIndex        =   1
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FF8080&
      Caption         =   "And Much, Much More"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6360
      TabIndex        =   16
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FF8080&
      Caption         =   "Delegation"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6360
      TabIndex        =   15
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FF8080&
      Caption         =   "Supervision"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6360
      TabIndex        =   14
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF8080&
      Caption         =   "Teamwork"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6360
      TabIndex        =   13
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF8080&
      Caption         =   "Assertiveness"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6360
      TabIndex        =   12
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF8080&
      Caption         =   "Communication"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6360
      TabIndex        =   11
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8080&
      Caption         =   "Leadership"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   6360
      TabIndex        =   10
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      Caption         =   "Mediation"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2040
      TabIndex        =   9
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "Conflict Resolution"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2040
      TabIndex        =   8
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Creativity"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Event Planning"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2040
      TabIndex        =   6
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "Budgeting"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Organization"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Time Management Skills"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form7.frx":0000
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   11655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Skills And Benefits"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMenu_Click()
Form7.Hide
Form2.Show
End Sub
