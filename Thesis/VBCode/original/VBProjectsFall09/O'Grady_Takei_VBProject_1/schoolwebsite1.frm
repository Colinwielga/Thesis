VERSION 5.00
Begin VB.Form frmhome 
   BackColor       =   &H80000002&
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   8880
      TabIndex        =   3
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdmajor 
      Caption         =   "View Class List!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4200
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
   End
   Begin VB.PictureBox picCSBSJU 
      Height          =   2415
      Left            =   360
      Picture         =   "school website1.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   1200
      Width           =   10815
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Welcome to School  web site"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "frmhome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This program allows the user to read SJU/CSB coursework into multiple arrays
' Written By John O'Grady and Yuzu Takei
' Written 10-15-09


Private Sub cmdmajor_Click()
    frmmajrlist.Show
    frmhome.Hide

End Sub

Private Sub cmdquit_Click()
'to quit
End
End Sub
