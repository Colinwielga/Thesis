VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF0000&
   Caption         =   "Main Page"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   5685
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C00000&
      Caption         =   "Start Here!"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8520
      MaskColor       =   &H000000C0&
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Caption         =   "Fun With CSB/SJU History!"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   1920
      TabIndex        =   0
      Top             =   7080
      Width           =   11535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
'Fun with CSB/SJU History!
'frmMain
'Audrey Gabe
'Written 3/12/09
'This project is to educate people in an entertaining way about CSB/SJU history while using different visual basic techniques.

frmMain.Hide
frmMenu.Show
End Sub


