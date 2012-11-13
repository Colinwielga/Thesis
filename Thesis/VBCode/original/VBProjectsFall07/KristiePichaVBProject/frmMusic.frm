VERSION 5.00
Begin VB.Form frmMusic 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Back to Menu"
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdChamberchoir 
      BackColor       =   &H0000C000&
      Caption         =   "Chamber Choir"
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdWomenschoir 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Women's Choir"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdConcertChoir 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Concert Choir"
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdBand 
      BackColor       =   &H0000C000&
      Caption         =   "Band"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBand_Click()
MsgBox "Kristie was in band for three years in 7th-9th grades."
End Sub

Private Sub cmdChamberchoir_Click()
MsgBox "Kristie was in Chamber Choir for one year in 12th grade."
End Sub

Private Sub cmdConcertChoir_Click()
MsgBox "Kristie was in Concert Choir for three years from 10th-12th grades."
End Sub

Private Sub cmdQuit_Click()
frmMusic.Hide
frmMenu.Show
End Sub

Private Sub cmdWomenschoir_Click()
MsgBox "Kristie was in Women's Choir for two years in 11th-12th grades."
End Sub

