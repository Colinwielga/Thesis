VERSION 5.00
Begin VB.Form frmSavePrincess2 
   BackColor       =   &H00000000&
   Caption         =   "Save the Princess!"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   8280
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSavePrincess2.frx":0000
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   3375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
   Begin VB.Image img6 
      Height          =   2085
      Left            =   600
      Picture         =   "frmSavePrincess2.frx":009C
      Top             =   3360
      Width           =   2220
   End
   Begin VB.Image img7 
      Height          =   2085
      Left            =   2760
      Picture         =   "frmSavePrincess2.frx":1322
      Top             =   3360
      Width           =   2220
   End
   Begin VB.Image img8 
      Height          =   2085
      Left            =   4920
      Picture         =   "frmSavePrincess2.frx":25F8
      Top             =   3360
      Width           =   2220
   End
   Begin VB.Image img9 
      Height          =   2085
      Left            =   7080
      Picture         =   "frmSavePrincess2.frx":38AD
      Top             =   3360
      Width           =   2220
   End
End
Attribute VB_Name = "frmSavePrincess2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'Save the princess
'the user has to click on a cave to take the user to another form
'Each cave will have something different to bring to the table

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub img6_Click()
frmCave12.Show
frmSavePrincess2.Hide

End Sub

Private Sub img7_Click()
frmCave22.Show
frmSavePrincess2.Hide
End Sub

Private Sub img8_Click()
frmCave32.Show
frmSavePrincess2.Hide
End Sub

Private Sub img9_Click()
frmCave42.Show
frmSavePrincess2.Hide
End Sub
