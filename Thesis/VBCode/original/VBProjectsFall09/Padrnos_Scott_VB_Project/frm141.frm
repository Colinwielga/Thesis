VERSION 5.00
Begin VB.Form frm141 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19185
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   19185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   17880
      TabIndex        =   11
      Top             =   9600
      Width           =   1215
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
      Height          =   1455
      Left            =   14880
      TabIndex        =   10
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "141 lb. Weight Class"
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
      Left            =   6720
      TabIndex        =   12
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblGoldschmidt2 
      Caption         =   $"frm141.frx":0000
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
      Left            =   15120
      TabIndex        =   9
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label lblKirscht2 
      Caption         =   $"frm141.frx":011F
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   10440
      TabIndex        =   8
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label lblDerik2 
      Caption         =   $"frm141.frx":0222
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
      Left            =   6840
      TabIndex        =   7
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label lblMorse2 
      Caption         =   $"frm141.frx":0342
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
      Left            =   4080
      TabIndex        =   6
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label lblMinga2 
      Caption         =   $"frm141.frx":0412
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label lblGoldschmidt 
      Caption         =   "Fr. Cody Goldschmidt"
      Height          =   255
      Left            =   15720
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblKirscht 
      Caption         =   "Fr. Charlie Kirscht"
      Height          =   255
      Left            =   11400
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblDerik 
      Caption         =   "So. Derik Gertken"
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblMorse 
      Caption         =   "Sr. Sam Morse"
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblMinga 
      Caption         =   "Jr. Myanganbayar Batsukh"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Image Image5 
      Height          =   2415
      Left            =   15000
      Picture         =   "frm141.frx":051B
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   2400
      Left            =   10200
      Picture         =   "frm141.frx":1AE81
      Top             =   2760
      Width           =   3600
   End
   Begin VB.Image Image3 
      Height          =   2640
      Left            =   7080
      Picture         =   "frm141.frx":1D489
      Top             =   2760
      Width           =   2220
   End
   Begin VB.Image Image2 
      Height          =   2595
      Left            =   4440
      Picture         =   "frm141.frx":1F00C
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   360
      Picture         =   "frm141.frx":1FF61
      Top             =   1320
      Width           =   3540
   End
End
Attribute VB_Name = "frm141"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
 frm141.Hide 'hides the 141 form
 frmCompetition.Show    'brings up the competition form
End Sub

Private Sub cmdQuit_Click()
'ends the program
    End
End Sub
