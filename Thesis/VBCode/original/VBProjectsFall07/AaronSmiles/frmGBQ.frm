VERSION 5.00
Begin VB.Form frmGBQ 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdHerc 
      Caption         =   "Herculean"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   4335
   End
   Begin VB.CommandButton cmdKOD 
      Caption         =   "Kingdom of Doom"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   4335
   End
   Begin VB.CommandButton cmdHS 
      Caption         =   "History Song"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CommandButton cmdBTS 
      Caption         =   "Behind The Sun"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmGBQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Playlist of selected songs from "the band the good the bad and the queen"
Private Sub cmdBTS_Click()
Shell ("explorer.exe M:\CS130\Try 3\08 behind the sun.mp3")
End Sub

Private Sub cmdexit_Click()
frmGBQ.Hide
End Sub

Private Sub cmdHerc_Click()
Shell ("explorer.exe M:\CS130\Try 3\05 Herculean.mp3")
End Sub

Private Sub cmdHS_Click()
Shell ("explorer.exe M:\CS130\Try 3\01 - history song.mp3")
End Sub

Private Sub cmdKOD_Click()
Shell ("explorer.exe M:\CS130\Try 3\04 Kingdom Of Doom.mp3")
End Sub

