VERSION 5.00
Begin VB.Form frmPictureMenu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "picture menu"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPicture 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit3 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      Height          =   495
      Left            =   3600
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdReturnEnglish 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to English Menu"
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdRefresh2 
      BackColor       =   &H00FF8080&
      Caption         =   "refresh files"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.FileListBox File2 
      Height          =   1065
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
   Begin VB.DirListBox Dir2 
      Height          =   1890
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "please type complete path to your 100 x 100 .bmp picture in text field below! "
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   3615
   End
End
Attribute VB_Name = "frmPictureMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim B As String

Private Sub cmdExit3_Click()
B = MsgBox("Are You Sure You Want To Exit?! This is a really cool program!", vbQuestion + vbYesNo, "Exit ?")
If B = 6 Then
    End
  Else
 frmLanguageMenu.Show
  End If
End Sub

Private Sub cmdRefresh2_Click()
    Dir2.Path = "M:"
    File2.Path = Dir2.Path
End Sub

Private Sub cmdReturnEnglish_Click()
        txtPicture.Text = PicPath
    frmSoundMenu.Hide
    frmEnglishMenu.Show
End Sub

