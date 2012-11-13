VERSION 5.00
Begin VB.Form frmReebokSeniors 
   BackColor       =   &H000040C0&
   Caption         =   "ReebokSeniors"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   Picture         =   "frmReebokSeniors.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00004000&
      Caption         =   "Load"
      Height          =   975
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   7215
      Left            =   8880
      ScaleHeight     =   7155
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   0
      Width           =   6255
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00C0C000&
      Caption         =   "Back To Reebok Home"
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000FFFF&
      Caption         =   "Back To Store Home"
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1455
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   855
      Left            =   120
      OleObjectBlob   =   "frmReebokSeniors.frx":BBB8A
      SourceDoc       =   "M:\CS130\AndrewEisingerVBproject\ESPN_-_National_Hockey_Night.mp3"
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmReebokSeniors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' AthleticStore
' ReebokSeniors
' Andrew Eisinger
' 3/23/09
'This program loads info into the picture box




Private Sub cmdBack_Click()
frmStoreHome.Show
frmReebokSeniors.Hide
End Sub

Private Sub cmdGoBack_Click()
frmReebok1.Show
frmReebokSeniors.Hide
End Sub



Private Sub cmdLoad_Click()
picResults.Print ("Whose kidding, Seniors don't play sports!")

End Sub

Private Sub cmdQuit_Click()
End
End Sub



        
