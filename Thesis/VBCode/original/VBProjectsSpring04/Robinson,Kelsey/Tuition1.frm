VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11040
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   11040
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go to the Next Slide"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   6135
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Average application fee...$50"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   9855
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Flight to tour east coast colleges...$400"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   9855
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Amount spent on cans of Mountain Dew to study for the SATs/ACTs...$9"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   9855
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Having a computer program decide where you REALLY want to go to college........................ PRICELESS"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deciding on a College
' Form 1 (Tuition1)
' Kelsey Robinson
' March 10th, 2004
'This program sorts through a list of colleges to determine which college is best for a perspective student with regards to tuition and distance from the Twin Cities.

Option Explicit
Dim College(1 To 20) As String, Tuition(1 To 20) As Single, Distance(1 To 20) As Single
Dim Specialties(1 To 20) As String
Dim CTR As Single


Private Sub cmdNext_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub



