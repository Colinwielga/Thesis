VERSION 5.00
Begin VB.Form frmHockeyStatistics 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8505
   FillColor       =   &H000000FF&
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   3360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoalies 
      Caption         =   "See Top 10 Goalies"
      Height          =   1215
      Left            =   5760
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdDefense 
      Caption         =   "See Top10 Defense"
      Height          =   1215
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdForwards 
      BackColor       =   &H000000FF&
      Caption         =   "See Top 10 Forwards"
      Height          =   1215
      Left            =   840
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblHockeyStats 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Hockey Statistics for the 2009-2010 NHL season"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   7185
   End
End
Attribute VB_Name = "frmHockeyStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Hockey Statistics
'Form Name: frmhockeystatistics
'Autor: Weston Lake
'Date Written: October 19, 2009
'Objective: To see who are the best players in the NHL so far this season based on their stats
Option Explicit

Private Sub cmdDefense_Click()
    Close #1
    frmDefense.Show
    frmHockeyStatistics.Hide
End Sub

Private Sub cmdForwards_Click()
    Close #1
    frmForwards.Show
    frmHockeyStatistics.Hide

End Sub

Private Sub cmdGoalies_Click()
 Close #1
 frmGoalie.Show
 frmHockeyStatistics.Hide
End Sub

Private Sub picResults_Click()
 picResults LoadPicture(App.Path & "\anderson.jpg")
End Sub
