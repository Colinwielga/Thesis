VERSION 5.00
Begin VB.Form welcome 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C0C0&
      Height          =   6015
      Left            =   480
      Picture         =   "welcome.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   960
      Width           =   9015
      Begin VB.CommandButton cmdquit 
         BackColor       =   &H0000FFFF&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdnext 
         BackColor       =   &H000000C0&
         Caption         =   "Click Here to Begin the Fun!"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6840
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Label lblfirst 
      BackColor       =   &H0000FFFF&
      Caption         =   "Welcome to  "
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdnext_Click()
    'this button enables the user to continue on to the next form
    welcome.Hide
    Login.Show
End Sub

Private Sub cmdquit_Click()
    End
End Sub
