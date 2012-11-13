VERSION 5.00
Begin VB.Form frm149 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19725
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   19725
   StartUpPosition =   3  'Windows Default
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
      Height          =   1455
      Left            =   17160
      TabIndex        =   6
      Top             =   8280
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   15600
      TabIndex        =   5
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "149 lb. Weight Class"
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
      Left            =   7800
      TabIndex        =   12
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblGlynn2 
      Caption         =   $"frm149.frx":0000
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
      Left            =   11640
      TabIndex        =   11
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label lblVaith2 
      Caption         =   $"frm149.frx":014C
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
      Left            =   8160
      TabIndex        =   10
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label lblDrew2 
      Caption         =   $"frm149.frx":0271
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4320
      TabIndex        =   9
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label lblGoldschmidt4 
      Caption         =   $"frm149.frx":0407
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   15720
      TabIndex        =   8
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label lblMinga4 
      Caption         =   $"frm149.frx":0526
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
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label lblGoldschmidt3 
      Caption         =   "Fr. Cody Goldschmidt"
      Height          =   255
      Left            =   16440
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblVaith 
      Caption         =   "Jr. John Paul Vaith"
      Height          =   255
      Left            =   9000
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblGlynn 
      Caption         =   "Fr. Kyle Glynn"
      Height          =   255
      Left            =   12720
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblDrew 
      Caption         =   "Jr. Drew Larson"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblMinga3 
      Caption         =   "Jr. Myanganbayar Batsukh"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Image Image5 
      Height          =   2415
      Left            =   15600
      Picture         =   "frm149.frx":062F
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   2805
      Left            =   8160
      Picture         =   "frm149.frx":1AF95
      Top             =   2640
      Width           =   3030
   End
   Begin VB.Image Image3 
      Height          =   2490
      Left            =   11520
      Picture         =   "frm149.frx":1DA80
      Top             =   2760
      Width           =   3810
   End
   Begin VB.Image Image2 
      Height          =   3285
      Left            =   4680
      Picture         =   "frm149.frx":1FF17
      Top             =   2280
      Width           =   2595
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   240
      Picture         =   "frm149.frx":258FA
      Top             =   1320
      Width           =   3540
   End
End
Attribute VB_Name = "frm149"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frm149.Hide 'hides the 149 form
frmCompetition.Show 'shows the cmopetition form
End Sub

Private Sub cmdQuit_Click()
'ends the program
    End
End Sub

