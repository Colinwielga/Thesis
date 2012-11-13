VERSION 5.00
Begin VB.Form frmOpener 
   BackColor       =   &H8000000D&
   Caption         =   "Continue"
   ClientHeight    =   11580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18090
   LinkTopic       =   "Form1"
   Picture         =   "VBProjectOpener.frx":0000
   ScaleHeight     =   11580
   ScaleWidth      =   18090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   1575
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CommandButton cmdProject 
      BackColor       =   &H0000FFFF&
      Caption         =   "Continue"
      Height          =   1575
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   2535
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   1215
      Left            =   5040
      OleObjectBlob   =   "VBProjectOpener.frx":FC8E2
      SourceDoc       =   "M:\CS130\AndrewEisingerVBproject\NBA_on_NBC_-_Live.mp3"
      TabIndex        =   2
      Top             =   7200
      Width           =   2295
   End
End
Attribute VB_Name = "frmOpener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
'Opener
'Andrew Eisinger
'on March 22
'This form will move to the next form unless quit


Private Sub cmdProject_Click()
'This project moves to the next form
Found = False
frmOpener.Hide
frmStoreHome.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

