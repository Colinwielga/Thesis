VERSION 5.00
Begin VB.Form frmGeography 
   BackColor       =   &H00FF00FF&
   Caption         =   "Geography"
   ClientHeight    =   8850
   ClientLeft      =   1470
   ClientTop       =   1215
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGeo 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   1
      Text            =   "Geography"
      Top             =   0
      Width           =   5655
   End
   Begin VB.PictureBox picJ4 
      Height          =   9015
      Left            =   -120
      Picture         =   "frmJ4.frx":0000
      ScaleHeight     =   8955
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   -120
      Width           =   12015
      Begin VB.CommandButton cmdScore 
         BackColor       =   &H00FF8080&
         Caption         =   "Click to View Your Current Score!"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FF8080&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8040
         Width           =   1815
      End
      Begin VB.CommandButton cmdReturn 
         BackColor       =   &H00FF8080&
         Caption         =   "Return To Topic Options"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7800
         Width           =   2715
      End
      Begin VB.CommandButton cmdG400 
         BackColor       =   &H00FF8080&
         Caption         =   "400"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3840
         Width           =   3135
      End
      Begin VB.CommandButton cmdG300 
         BackColor       =   &H00FF8080&
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3720
         Width           =   3135
      End
      Begin VB.CommandButton cmdG200 
         BackColor       =   &H00FF8080&
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CommandButton cmdG100 
         BackColor       =   &H00FF8080&
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmGeography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim G100 As String, G200 As String, G300 As String, G400 As String
'Jeopardy.(Jeopardy.vbp)
'Form name: Geography; Form caption: Jeopardy
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: This is the Geography form. Like before, it displays questions pertaining
'                to geography.

Private Sub cmdG400_Click()
G400 = InputBox("Where state is Mount Rushmore located in? Please answer in all lower case letters", , G400)
    If G400 = "south dakota" Then
        MsgBox "You are Correct!  Great Work!", , "Answer"
        Sum = Sum + 400
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 400
    End If
    cmdG400.Visible = False
        If cmdG100.Visible = False And cmdG200.Visible = False And cmdG300.Visible = False And cmdG400.Visible = False Then
            frmGeography.Hide
            frmTopics.Show
        End If
End Sub

Private Sub cmdG200_Click()
G200 = InputBox("What state is known as The Big Apple? Please answer in all lower case letters", , G200)
    If G200 = "new york" Then
        MsgBox "You are Correct!  Great Work!", , "Answer"
        Sum = Sum + 200
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 200
    End If
    cmdG200.Visible = False
        If cmdG100.Visible = False And cmdG200.Visible = False And cmdG300.Visible = False And cmdG400.Visible = False Then
            frmGeography.Hide
            frmTopics.Show
        End If
End Sub

Private Sub cmdG100_Click()
G100 = InputBox("What is the capital of California? Please answer in all lower case letters", , G100)
    If G100 = "sacramento" Then
        MsgBox "You are Correct!  Good Answer!", , "Answer"
        Sum = Sum + 100
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 100
    End If
    cmdG100.Visible = False
        If cmdG100.Visible = False And cmdG200.Visible = False And cmdG300.Visible = False And cmdG400.Visible = False Then
            frmGeography.Hide
            frmTopics.Show
        End If
        
End Sub

Private Sub cmdG300_Click()
G300 = InputBox("What state is the Yellowstone National Park located in? Please answer in all lower case letters", , G300)
    If G300 = "wyoming" Then
        MsgBox "You are Correct!  Terrific!", , "Answer"
        Sum = Sum + 300
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 300
    End If
    cmdG300.Visible = False
        If cmdG100.Visible = False And cmdG200.Visible = False And cmdG300.Visible = False And cmdG400.Visible = False Then
            frmGeography.Hide
            frmTopics.Show
        End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmGeography.Hide
    frmTopics.Show
End Sub

Private Sub cmdScore_Click()
    MsgBox "Your score is " & Sum & ". Nice work!"
End Sub

