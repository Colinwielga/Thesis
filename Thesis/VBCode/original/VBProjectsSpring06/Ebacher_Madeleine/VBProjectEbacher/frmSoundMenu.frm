VERSION 5.00
Begin VB.Form frmSoundMenu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Sound Menu"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSound 
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh list"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdSetSound 
      Caption         =   "Save Sound"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   5880
      Width           =   3975
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblMyProgram 
      BackColor       =   &H00FFC0C0&
      Caption         =   "VB Alarm Clock - Madeleine Ebacher"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label lblBrowsin 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Please find a .wav sound you would like to play:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmSoundMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB Alarm Clock (project1.vbp)
'"Sound Menu" (frmSoundMenu.frm)
'designed by: Madeleine Ebacher
'3/24/06
'This menu allows the user to choose a .wav sound to play as a notification.

Option Explicit
Dim SoundPath As String

Private Sub cmdRefresh_Click()
Dir1.Path = "M:"
File1.Path = Dir1.Path
End Sub

Private Sub cmdSetSound_Click()
    txtSound.Text = SoundPath
    frmSoundMenu.Hide
    frmEnglishMenu.Show
        
End Sub


