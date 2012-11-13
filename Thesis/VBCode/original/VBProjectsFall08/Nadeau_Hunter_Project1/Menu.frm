VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10035
   ClientLeft      =   2595
   ClientTop       =   1815
   ClientWidth     =   15240
   FillColor       =   &H00C00000&
   BeginProperty Font 
      Name            =   "Perpetua"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdBiblio 
      Caption         =   "Bibliography"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11040
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12120
      TabIndex        =   3
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdLibrary 
      Caption         =   "Library"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8040
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdAstro 
      Caption         =   "Astrology"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdBio 
      BackColor       =   &H0080FFFF&
      Caption         =   "Biography"
      DisabledPicture =   "Menu.frx":5D120
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      MaskColor       =   &H0080FFFF&
      Picture         =   "Menu.frx":69667
      TabIndex        =   0
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Caption         =   "eccentric bachelor, deceased"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   5640
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Natural Philosophy of Reuben Brown"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1200
      TabIndex        =   5
      Top             =   5160
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CSCI 130 01A Project                                  By Zach Hunter and Nik Nadeau"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   9000
      Width           =   9255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Reuben Brown
'Form name: Menu
'Author: Nik Nadeau and Zach Hunter
'Date: Nov. 4, 2008
'This form is the main menu. The below subroutines allow user to move to different forms.
 
Private Sub cmdAstro_Click()
Form1.Hide
Astrology.Show
End Sub

Private Sub cmdBiblio_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub cmdBio_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub cmdLibrary_Click()
Form1.Hide
Form4.Show
End Sub

Private Sub cmdPhilo_Click()
Form1.Hide
Form5.Show
End Sub

Private Sub cmdTarot_Click()
Form1.Hide
Form4.Show
End Sub
