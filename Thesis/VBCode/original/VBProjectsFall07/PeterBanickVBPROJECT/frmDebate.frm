VERSION 5.00
Begin VB.Form frmDebate 
   Caption         =   "The Issues ..."
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10890
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdReturnMenu 
      BackColor       =   &H000000FF&
      DisabledPicture =   "frmDebate.frx":0000
      Height          =   1215
      Left            =   12960
      Picture         =   "frmDebate.frx":7DDC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CommandButton cmdReadDebate 
      BackColor       =   &H80000007&
      Height          =   1935
      Left            =   360
      Picture         =   "frmDebate.frx":F721
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   3015
   End
   Begin VB.PictureBox picResultsDebate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   10935
      Left            =   0
      Picture         =   "frmDebate.frx":19F4D
      ScaleHeight     =   10905
      ScaleWidth      =   15225
      TabIndex        =   0
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmDebate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReadDebate_Click()
    'reads in text from file \debate.txt for user to read further about Rose and learn about his controversy
    Dim Debate As String
    picResultsDebate.Cls
    Open App.Path & "\debate.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Debate
        picResultsDebate.Print Debate
    Loop
    Close #1
    cmdReadDebate.Enabled = False
End Sub

Private Sub cmdReturnMenu_Click()
    'returns user bio page, removes debate page from visibility
    cmdReadDebate.Enabled = True
    frmDebate.Hide
End Sub
