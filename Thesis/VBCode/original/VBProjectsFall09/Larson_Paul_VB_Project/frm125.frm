VERSION 5.00
Begin VB.Form frm125 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   16155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Redirect Back To Weight Classes"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   14280
      TabIndex        =   4
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   14640
      TabIndex        =   0
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "125 lb. Weight Class"
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
      Left            =   5520
      TabIndex        =   8
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblSeck2 
      Caption         =   $"frm125.frx":0000
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   10800
      TabIndex        =   7
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label lblScott2 
      Caption         =   $"frm125.frx":00D1
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5280
      TabIndex        =   6
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label lblHenle2 
      Caption         =   $"frm125.frx":01A4
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
      Left            =   360
      TabIndex        =   5
      Top             =   5040
      Width           =   3975
   End
   Begin VB.Image Image3 
      Height          =   3000
      Left            =   11160
      Picture         =   "frm125.frx":02C7
      Top             =   1200
      Width           =   2310
   End
   Begin VB.Image Image2 
      Height          =   2430
      Left            =   5520
      Picture         =   "frm125.frx":16D89
      Top             =   2160
      Width           =   3075
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   840
      Picture         =   "frm125.frx":18D30
      Top             =   1080
      Width           =   2685
   End
   Begin VB.Label lblSeck 
      Caption         =   "Fr. Trenton Seck"
      Height          =   255
      Left            =   11400
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblScott 
      Caption         =   "So. Scott Padrnos"
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblHenle 
      Caption         =   "So. Chad Henle"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frm125"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
    frm125.Hide 'Hides the form frm125
    frmCompetition.Show 'Pulls up the form frmCompetition
End Sub


'Quits the program
Private Sub cmdQuit_Click()
    End
End Sub


