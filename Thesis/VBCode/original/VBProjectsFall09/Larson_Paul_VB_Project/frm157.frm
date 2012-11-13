VERSION 5.00
Begin VB.Form frm157 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   11595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20235
   LinkTopic       =   "Form1"
   ScaleHeight     =   11595
   ScaleWidth      =   20235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   1320
      TabIndex        =   13
      Top             =   10320
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to weight classes"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   16560
      TabIndex        =   12
      Top             =   8760
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "157 lb. Weight Class"
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
      Left            =   7320
      TabIndex        =   14
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblGlynn4 
      Caption         =   $"frm157.frx":0000
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
      Index           =   1
      Left            =   16200
      TabIndex        =   11
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label lblStevermer2 
      Caption         =   $"frm157.frx":014C
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   13080
      TabIndex        =   10
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label lblAnderson2 
      Caption         =   $"frm157.frx":02AB
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   0
      Left            =   10320
      TabIndex        =   9
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label lblVaith4 
      Caption         =   $"frm157.frx":0385
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   4080
      TabIndex        =   8
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label lblDrew4 
      Caption         =   $"frm157.frx":04AA
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Index           =   1
      Left            =   7200
      TabIndex        =   7
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label lblBaarson2 
      Caption         =   $"frm157.frx":0640
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label lblGlynn3 
      Caption         =   "Fr. Kyle Glynn"
      Height          =   255
      Left            =   17160
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblStevermer 
      Caption         =   "Fr. Chris Stevermer"
      Height          =   255
      Left            =   13320
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblAnderson 
      Caption         =   "Sr. Ben Anderson"
      Height          =   255
      Left            =   10680
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblDrew3 
      Caption         =   "Jr. Drew Larson"
      Height          =   255
      Left            =   7920
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblVaith3 
      Caption         =   "Jr. John Paul Vaith"
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblBaarson 
      Caption         =   "Jr. Matt Baarson"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   2250
      Left            =   12840
      Picture         =   "frm157.frx":07FA
      Top             =   2640
      Width           =   2430
   End
   Begin VB.Image Image6 
      Height          =   3285
      Left            =   7320
      Picture         =   "frm157.frx":2493
      Top             =   2640
      Width           =   2595
   End
   Begin VB.Image Image5 
      Height          =   2490
      Left            =   15840
      Picture         =   "frm157.frx":7E76
      Top             =   1560
      Width           =   3810
   End
   Begin VB.Image Image3 
      Height          =   2880
      Left            =   10440
      Picture         =   "frm157.frx":A30D
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Image Image2 
      Height          =   2805
      Left            =   3960
      Picture         =   "frm157.frx":B4EF
      Top             =   2640
      Width           =   3030
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   240
      Picture         =   "frm157.frx":DFDA
      Top             =   1680
      Width           =   3405
   End
End
Attribute VB_Name = "frm157"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frm157.Hide ' hides the 157 form
frmCompetition.Show 'brings up the competition form
End Sub

Private Sub cmdQuit_Click()
'ends the program
    End
End Sub
