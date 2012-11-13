VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H80000007&
   Caption         =   "Welcome"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter Stats"
      Height          =   615
      Left            =   2520
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      Begin VB.PictureBox PicNFCTeams 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   4845
         Left            =   720
         Picture         =   "frmWelcome.frx":0000
         ScaleHeight     =   4785
         ScaleWidth      =   4965
         TabIndex        =   0
         Top             =   0
         Width           =   5025
      End
      Begin VB.Shape Shape1 
         FillStyle       =   2  'Horizontal Line
         Height          =   5175
         Left            =   0
         Top             =   0
         Width           =   6135
      End
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H8000000D&
      Caption         =   "Quit"
      Height          =   615
      Left            =   4320
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton CmdWelcome 
      BackColor       =   &H000000FF&
      Caption         =   "Enter NFC North Teams"
      Height          =   615
      Left            =   840
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: Brandon Kasper
'Written 10/19/2009
'This project displays information for a user about the NFC North
'A user can enter multiple forms and read files from arrays, view pictures
'be interactive with players, and enter inputs.

Private Sub cmdquit_Click()
    End 'closes program
End Sub


Private Sub cmdStats_Click()
    Person = InputBox("What's your name?", "Welcome!")
    frmWelcome.Hide 'hides Welcome page from user
    FrmOD.Show 'shows stats page to user
    MsgBox "This is the NFC North Stats, " & Person & ".", , "Football"
End Sub

Private Sub CmdWelcome_Click()
    Person = InputBox("What's your name?", "Welcome!")
    frmWelcome.Hide 'hides Welcome page from user
    frmTeams.Show 'shows teams page to user
    MsgBox "This is the NFC North, " & Person & ".", , "Football"
End Sub

