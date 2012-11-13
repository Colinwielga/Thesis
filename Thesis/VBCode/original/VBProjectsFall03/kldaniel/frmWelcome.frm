VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Welcome to Bible Resource"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLook 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Verse Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   3855
   End
   Begin VB.CommandButton cmdToConc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Bible Concordance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   3855
   End
   Begin VB.PictureBox pbxBible 
      Height          =   6135
      Left            =   360
      Picture         =   "frmWelcome.frx":0000
      ScaleHeight     =   6075
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Welcome to Bible Resource"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   10575
   End
   Begin VB.Label lblPick 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pick one of the following Options:"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   975
      Left            =   7080
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLook_Click()
frmWelcome.Hide
frmVerseSearch.Show
End Sub


Private Sub cmdToConc_Click()
frmWelcome.Hide
frmConc.Show
End Sub

Private Sub Form_Load()
Dim strTemp, strTemp2 As String
Dim k, t As Integer
Dim End_Of_File As Boolean
strPath = "N:\CS130\handin\kldaniel\"
Open strPath & "Romans1.txt" For Input As #1
End_Of_File = False
Input #1, strBook
Input #1, strChapter(1)
For k = 1 To 32
    Line Input #1, strTemp
    strVerse(1, k) = strTemp
Next k
Close #1
End Sub


