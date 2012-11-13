VERSION 5.00
Begin VB.Form frmBio 
   Caption         =   "Pete Rose's Life"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdDebate 
      BackColor       =   &H000000FF&
      Height          =   1935
      Left            =   6960
      Picture         =   "frmBio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   3015
   End
   Begin VB.PictureBox picResultsBio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   10575
      Left            =   0
      Picture         =   "frmBio.frx":A305
      ScaleHeight     =   10545
      ScaleWidth      =   15225
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.CommandButton cmdReturnMenu 
         BackColor       =   &H000000FF&
         DisabledPicture =   "frmBio.frx":37F38
         Height          =   1215
         Left            =   13080
         Picture         =   "frmBio.frx":3FD14
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6840
         Width           =   1935
      End
      Begin VB.CommandButton cmdReadBio 
         BackColor       =   &H000000FF&
         Height          =   1935
         Left            =   12120
         Picture         =   "frmBio.frx":47659
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   8400
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDebate_Click()
    'brings up formDebate to read further
    cmdReadBio.Enabled = True
    frmDebate.Show
End Sub

Private Sub cmdReadBio_Click()
    'reads in file \bio.txt to read about Pete Rose's life and background
    Dim Bio As String
    picResultsBio.Cls
    Open App.Path & "\bio.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Bio
        picResultsBio.Print Bio
    Loop
    Close #1
    cmdReadBio.Enabled = False
    cmdDebate.Enabled = True
End Sub


Private Sub cmdReturnMenu_Click()
    'returns user to menu page, removes bio page from visibility
    cmdReadBio.Enabled = True
    frmBio.Hide
    frmMenuPage.Show
End Sub

