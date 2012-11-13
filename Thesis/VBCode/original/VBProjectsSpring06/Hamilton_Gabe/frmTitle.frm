VERSION 5.00
Begin VB.Form frmTitle 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Title"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   DrawMode        =   16  'Merge Pen
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Gabe Hamilton"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblTrivia 
      BackStyle       =   0  'Transparent
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      ToolTipText     =   "Click Here For Trivia"
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblBio 
      BackStyle       =   0  'Transparent
      Caption         =   "Biographies"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      ToolTipText     =   "Click Here For Biographies"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "World Rankings"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      ToolTipText     =   "Click Here For World Rankings"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   5520
      Left            =   0
      Picture         =   "frmTitle.frx":0000
      Top             =   0
      Width           =   7350
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   600
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Title acts as a home page and uses labels to navigate to other forms

Private Sub lblBio_Click()
    frmTitle.Hide
    frmBiographies.Show
End Sub

Private Sub lblExit_Click()
    End
End Sub

Private Sub lblRank_Click()
    frmTitle.Hide
    frmWorldRankings.Show
End Sub

Private Sub lblTrivia_Click()
    frmTitle.Hide
    frmTrivia.Show
End Sub
