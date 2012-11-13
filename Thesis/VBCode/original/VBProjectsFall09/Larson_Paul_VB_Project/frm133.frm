VERSION 5.00
Begin VB.Form frm133 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16380
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   16380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Weight Classes"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   14400
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   14640
      TabIndex        =   0
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "133 lb. Weight Class"
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
      Left            =   5760
      TabIndex        =   10
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblBoosalis2 
      Caption         =   $"frm133.frx":0000
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
      Left            =   11520
      TabIndex        =   9
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblManny2 
      Caption         =   $"frm133.frx":012D
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
      Left            =   7200
      TabIndex        =   8
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label lblLaine2 
      Caption         =   $"frm133.frx":0203
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
      Left            =   3960
      TabIndex        =   7
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label lblMogi2 
      Caption         =   $"frm133.frx":02C7
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
      Left            =   480
      TabIndex        =   6
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Image Image4 
      Height          =   3375
      Left            =   11520
      Picture         =   "frm133.frx":0402
      Top             =   960
      Width           =   2250
   End
   Begin VB.Image Image3 
      Height          =   2400
      Left            =   7320
      Picture         =   "frm133.frx":2E8A
      Top             =   1920
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   2700
      Left            =   4200
      Picture         =   "frm133.frx":7C27
      Top             =   1920
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   360
      Picture         =   "frm133.frx":192D9
      Top             =   960
      Width           =   3090
   End
   Begin VB.Label lblBoosalis 
      Caption         =   "Fr. Alex Boosalis"
      Height          =   255
      Left            =   11760
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblManny 
      Caption         =   "Jr. Manny Livingstone"
      Height          =   255
      Left            =   7440
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLaine 
      Caption         =   "Sr. Matt Laine"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblMogi 
      Caption         =   "Sr. Mogi Baatar"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frm133"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frm133.Hide 'hide form 133
frmCompetition.Show 'pull up competition form

End Sub

Private Sub cmdQuit_Click()
    End 'close
End Sub

