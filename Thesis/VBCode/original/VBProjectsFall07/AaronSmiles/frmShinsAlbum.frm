VERSION 5.00
Begin VB.Form frmShinsAlbum 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdNweSlang 
      Caption         =   "New Slang"
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
      TabIndex        =   3
      Top             =   2640
      Width           =   4455
   End
   Begin VB.CommandButton cmdAustralia 
      Caption         =   "Australia"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton cmdTurnOnME 
      Caption         =   "Turn On Me"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton cmdPhantomLimb 
      Caption         =   "Phantom Limb"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmShinsAlbum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAustralia_Click()
    Shell ("explorer.exe M:\CS130\Try 3\02. Australia.mp3")
End Sub

Private Sub cmdNweSlang_Click()
    Shell ("explorer.exe M:\CS130\Try 3\Shins, The - New Slang.mp3")
End Sub

Private Sub cmdPhantomLimb_Click()
    Shell ("explorer.exe M:\CS130\Try 3\Shins, The - New Slang.mp3")
End Sub

Private Sub cmdTurnOnME_Click()
    Shell ("explorer.exe M:\CS130\Try 3\07. Turn on Me.mp3")
End Sub

