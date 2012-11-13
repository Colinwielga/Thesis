VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H00FF0000&
   Caption         =   "History"
   ClientHeight    =   7035
   ClientLeft      =   2100
   ClientTop       =   1365
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHistory 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      TabIndex        =   1
      Text            =   "History"
      Top             =   0
      Width           =   3975
   End
   Begin VB.PictureBox picJ8 
      Height          =   7095
      Left            =   0
      Picture         =   "frmJ8.frx":0000
      ScaleHeight     =   7035
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton cmdScore 
         BackColor       =   &H00FF80FF&
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5400
         Width           =   2175
      End
      Begin VB.CommandButton cmdTopics 
         BackColor       =   &H00FF80FF&
         Caption         =   "Return to Topic Options"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5640
         Width           =   2175
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FF80FF&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdH400 
         BackColor       =   &H00FF80FF&
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
         Height          =   1245
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton cmdH300 
         BackColor       =   &H00FF80FF&
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
         Height          =   1245
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton cmdH200 
         BackColor       =   &H00FF80FF&
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CommandButton cmdH100 
         BackColor       =   &H00FF80FF&
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
         Height          =   1245
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim H100 As String, H200 As Integer, H300 As Integer, H400 As String
'Jeopardy.(Jeopardy.vbp)
'Form name: History; Form caption: Jeopardy
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: This is the history section of our program. Again, users will answer question
'                based on history.

Private Sub cmdH100_Click()
H100 = InputBox("Who was the first president of the United States? Please answer in all lower case letters", "H100")
    If H100 = "george washington" Then
        MsgBox "You are Correct!  Terrific!", , "Answer"
        Sum = Sum + 100                               ' This is the overall Total count(Sum)
    Else                                              ' It will add 100 if answer is correct
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 100                               ' This will subtract 100 if answer is
    End If                                            ' incorrect.
        cmdH100.Visible = False
    If cmdH100.Visible = False And cmdH200.Visible = False And cmdH300.Visible = False And cmdH400.Visible = False Then
        frmHistory.Hide                               ' This portion of the code makes it so
        frmTopics.Show                                ' if all questions have been answered,
    End If                                            ' it will return to the topics form.
End Sub

Private Sub cmdH200_Click()
H200 = InputBox("What year was the Declaration of Independence signed?", "H200")
    If H200 = 1776 Then
        MsgBox "You are Correct!  Way to Go!", , "Answer"
        Sum = Sum + 200
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 200
    End If
    cmdH200.Visible = False
    If cmdH100.Visible = False And cmdH200.Visible = False And cmdH300.Visible = False And cmdH400.Visible = False Then
        frmHistory.Hide
        frmTopics.Show
    End If
End Sub

Private Sub cmdH300_Click()
H300 = InputBox("How many original US colonies were there?", "H300")
    If H300 = "13" Then
        MsgBox "You are Correct!  Good Job!", , "Answer"
        Sum = Sum + 300
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 300
    End If
    cmdH300.Visible = False
    If cmdH100.Visible = False And cmdH200.Visible = False And cmdH300.Visible = False And cmdH400.Visible = False Then
        frmHistory.Hide
        frmTopics.Show
    End If
End Sub

Private Sub cmdH400_Click()
H400 = InputBox("At what theatre was President Lincoln shot at? Please enter in all lower case letters", "H400")
    If H400 = "ford" Then
        MsgBox "You are Correct!  Great Answer!", , "Answer"
        Sum = Sum + 400
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 400
    End If
    cmdH400.Visible = False
    If cmdH100.Visible = False And cmdH200.Visible = False And cmdH300.Visible = False And cmdH400.Visible = False Then
        frmHistory.Hide
        frmTopics.Show
    End If
End Sub


Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdScore_Click()
    MsgBox "Your score is " & Sum & ". Nice work!"
End Sub

Private Sub cmdTopics_Click()
    frmHistory.Hide
    frmTopics.Show
End Sub
