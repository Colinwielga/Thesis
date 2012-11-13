VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H0080FF80&
   Caption         =   "Welcome to Binary Review for Computer Science non-majors"
   ClientHeight    =   12135
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   12135
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbYear 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmWelcome.frx":0000
      Left            =   3720
      List            =   "frmWelcome.frx":0010
      TabIndex        =   7
      Top             =   5760
      Width           =   7095
   End
   Begin VB.TextBox txtLastScore 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   3
      Top             =   7080
      Width           =   2655
   End
   Begin VB.TextBox txtYourName 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5880
      TabIndex        =   2
      Top             =   3840
      Width           =   6615
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Start Practice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   6495
   End
   Begin VB.Image Image2 
      Height          =   2250
      Left            =   840
      Picture         =   "frmWelcome.frx":003A
      Top             =   2640
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   9480
      Picture         =   "frmWelcome.frx":DD92
      Top             =   9000
      Width           =   2310
   End
   Begin VB.Label lblLastScore 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter the Grade Your Got On Your Last Binary Quiz/Exam"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   7440
      Width           =   6135
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Your Year in School Here"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Your Name Here"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Welcome to Binary Practice!!"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   12855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   6  'Cross
      Height          =   975
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   6  'Cross
      Height          =   1215
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   6  'Cross
      Height          =   1095
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   7575
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
frmStudyGuide.Visible = True
frmAnswer.Visible = False
frmWelcome.Visible = False
YourName = txtYourName.Text
'Year = cmbYear.ComboBox
LastScore = txtLastScore.Text
End Sub

