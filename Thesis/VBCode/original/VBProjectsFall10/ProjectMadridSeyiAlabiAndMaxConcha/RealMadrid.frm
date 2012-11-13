VERSION 5.00
Begin VB.Form OpenPage 
   Caption         =   "Form1"
   ClientHeight    =   12180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   12180
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChange 
      Caption         =   "Next Page"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   0
      Top             =   10560
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   12000
      Left            =   0
      Picture         =   "Real Madrid.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13920
   End
End
Attribute VB_Name = "OpenPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdChange_Click()
Information.Show
Me.Hide
Form1.Hide
Statistics.Hide
PlayersStat.Hide
End Sub
