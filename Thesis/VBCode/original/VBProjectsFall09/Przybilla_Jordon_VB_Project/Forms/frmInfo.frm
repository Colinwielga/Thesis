VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H000040C0&
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfo 
      Caption         =   "See Program Information"
      Height          =   1575
      Left            =   7800
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.PictureBox picInfo 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   7155
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Go to Home Page"
      Height          =   1575
      Left            =   7800
      TabIndex        =   2
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1575
      Left            =   7800
      TabIndex        =   1
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   120
      Picture         =   "frmInfo.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Terms'
'Authors: Jordon Przybilla'
'Date Written: October 15, 2009
'this form will list the sources that i used to complete this project

Option Explicit

Private Sub cmdHome_Click()
'this button will take the user back to the home page
frmInfo.Hide
frmHome.Show


End Sub

Private Sub cmdInfo_Click()
'this button will display the array containing information for the program

Dim r As Integer
picInfo.Visible = True
picInfo.Cls

For r = 1 To Ctr
    picInfo.Print info(r)
    picInfo.Print
Next r



End Sub

Private Sub cmdQuit_Click()
End
End Sub

