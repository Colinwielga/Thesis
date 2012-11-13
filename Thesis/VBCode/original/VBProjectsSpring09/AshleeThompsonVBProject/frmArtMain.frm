VERSION 5.00
Begin VB.Form frmArtMain 
   BackColor       =   &H00808000&
   Caption         =   "Artwork"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   6
      Top             =   8400
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   4920
      Picture         =   "frmArtMain.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   4920
      Picture         =   "frmArtMain.frx":B908
      ScaleHeight     =   2955
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   3480
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   4920
      Picture         =   "frmArtMain.frx":D7AC
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   6960
      Width           =   3255
   End
   Begin VB.CommandButton cmdSketch 
      Caption         =   "Sketch Book"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton cmdPhoto 
      Caption         =   "Photography"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmdPaintings 
      Caption         =   "Paintings"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "frmArtMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The Artist's Multimedia Portfolio
'frmArtMain
'Ashley Thompson
'Friday March 20, 2009
'This form allows the user to navigate to other forms containing artwork using the .Hide and .Show functions
'It also allows the user to go back to the main menu in the same way

Private Sub cmdMain_Click()
frmArtMain.Hide
frmMain.Show
End Sub

Private Sub cmdPaintings_Click()
frmArt.Show
frmArtMain.Hide
End Sub

Private Sub cmdPhoto_Click()
frmArtMain.Hide
frmPhoto.Show
End Sub

Private Sub cmdSketch_Click()
frmArtMain.Hide
frmSketch.Show
End Sub
