VERSION 5.00
Begin VB.Form frm174 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16410
   BeginProperty Font 
      Name            =   "Viner Hand ITC"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   16410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to weight classes"
      Height          =   2175
      Left            =   14280
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14640
      TabIndex        =   6
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "174 lb. Weight Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblHagen2 
      Caption         =   $"frm174.frx":0000
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   10560
      TabIndex        =   5
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label lblPfarr2 
      Caption         =   $"frm174.frx":015F
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5640
      TabIndex        =   4
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label lblHagen 
      Caption         =   "Fr. Mitch Hagen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblPfarr 
      Caption         =   "So. Matt Pfarr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblRaygor 
      Caption         =   "Jr. Dustin Raygor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblRaygor2 
      Caption         =   $"frm174.frx":023A
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   960
      TabIndex        =   0
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   3120
      Left            =   1080
      Picture         =   "frm174.frx":0350
      Top             =   840
      Width           =   2190
   End
   Begin VB.Image Image2 
      Height          =   2190
      Left            =   5640
      Picture         =   "frm174.frx":1E8D
      Top             =   1920
      Width           =   3045
   End
   Begin VB.Image Image1 
      Height          =   2610
      Left            =   10920
      Picture         =   "frm174.frx":3EF2
      Top             =   1080
      Width           =   2115
   End
End
Attribute VB_Name = "frm174"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frm174.Hide 'hides the 174 form
frmCompetition.Show 'brings up the Competition form
End Sub

Private Sub cmdQuit_Click()
'ends the program
End
End Sub
