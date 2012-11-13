VERSION 5.00
Begin VB.Form frmSports 
   BackColor       =   &H00FF8080&
   Caption         =   "Sports"
   ClientHeight    =   7800
   ClientLeft      =   1935
   ClientTop       =   1680
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   7815
      Left            =   0
      Picture         =   "frmJ5.frx":0000
      ScaleHeight     =   7755
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton cmdScore 
         BackColor       =   &H00FFFF80&
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
         Height          =   1455
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FFFF80&
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
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6960
         Width           =   2055
      End
      Begin VB.CommandButton cmdReturn 
         BackColor       =   &H00FFFF80&
         Caption         =   "Return to Topic Options"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6960
         Width           =   2775
      End
      Begin VB.TextBox txtSports 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   3600
         TabIndex        =   5
         Text            =   "Sports"
         Top             =   120
         Width           =   3495
      End
      Begin VB.CommandButton cmdS400 
         BackColor       =   &H00FFFF80&
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
         Height          =   1575
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   3015
      End
      Begin VB.CommandButton cmdS300 
         BackColor       =   &H00FFFF80&
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
         Height          =   1575
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4560
         Width           =   3015
      End
      Begin VB.CommandButton cmdS200 
         BackColor       =   &H00FFFF80&
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
         Height          =   1575
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2280
         Width           =   3015
      End
      Begin VB.CommandButton cmdS100 
         BackColor       =   &H00FFFF80&
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
         Height          =   1575
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2280
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmSports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim S100 As String, S200 As String, S300 As Integer, S400 As Integer
'Jeopardy.(Jeopardy.vbp)
'Form name: Sports; Form caption: Jeopardy
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: This form is used as the sports category. Once again, there are four
'                questions that have values ranging from 100 to 400.

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmSports.Hide
    frmTopics.Show
End Sub

Private Sub cmdS100_Click()
S100 = InputBox("What sport is played on the ice with a puck? (Please enter in all lower case letters)", "S100")
    If S100 = "hockey" Then
        MsgBox "You are Correct!  Great Work!", , "Answer"
        Sum = Sum + 100
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 100
    End If
    cmdS100.Visible = False
        If cmdS100.Visible = False And cmdS200.Visible = False And cmdS300.Visible = False And cmdS400.Visible = False Then
            frmSports.Hide
            frmTopics.Show
        End If
End Sub

Private Sub cmdS200_Click()
S200 = InputBox("What is the name of the professional Minnesota basketball team?(Please enter in all lower case letters)", "S200")
    If S200 = "timberwolves" Then
        MsgBox "You are Correct!  Good Job!", , "Answer"
        Sum = Sum + 200
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 200
    End If
    cmdS200.Visible = False
        If cmdS100.Visible = False And cmdS200.Visible = False And cmdS300.Visible = False And cmdS400.Visible = False Then
            frmSports.Hide
            frmTopics.Show
        End If
End Sub

Private Sub cmdS300_Click()
S300 = InputBox("How many points is a touchdown worth?", "S300")
    If S300 = 6 Then
        MsgBox "You are Correct!  Wonderful!", , "Answer"
        Sum = Sum + 300
    Else
        MsgBox "Wrong Answer.", , " Wrong Answer"
        Sum = Sum - 300
    End If
    cmdS300.Visible = False
        If cmdS100.Visible = False And cmdS200.Visible = False And cmdS300.Visible = False And cmdS400.Visible = False Then
            frmSports.Hide
            frmTopics.Show
        End If
End Sub

Private Sub cmdS400_Click()
S400 = InputBox("How many World Series have the Minnesota Twins won?", "S400")
    If S400 = 2 Then
        MsgBox "You are Correct!  Way to Go!", , "Answer"
        Sum = Sum + 400
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 400
    End If
    cmdS400.Visible = False
        If cmdS100.Visible = False And cmdS200.Visible = False And cmdS300.Visible = False And cmdS400.Visible = False Then
            frmSports.Hide
            frmTopics.Show
        End If
End Sub

Private Sub cmdScore_Click()
    MsgBox "Your score is " & Sum & ". Nice work!"
End Sub
