VERSION 5.00
Begin VB.Form frmMath 
   Caption         =   "Math"
   ClientHeight    =   8625
   ClientLeft      =   1785
   ClientTop       =   1680
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMath 
      BackColor       =   &H00C0C000&
      Height          =   8415
      Left            =   0
      Picture         =   "frmJ3.frx":0000
      ScaleHeight     =   8355
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   120
      Width           =   11535
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
         Height          =   1095
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7200
         Width           =   2895
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FF80FF&
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7440
         Width           =   2415
      End
      Begin VB.CommandButton cmdReturnJ2 
         BackColor       =   &H00FF80FF&
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
         Height          =   735
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7440
         Width           =   2415
      End
      Begin VB.CommandButton cmdM400 
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
         Height          =   1695
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5040
         Width           =   3015
      End
      Begin VB.CommandButton cmdM300 
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
         Height          =   1695
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5040
         Width           =   3015
      End
      Begin VB.CommandButton cmdM200 
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
         Height          =   1695
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2280
         Width           =   3015
      End
      Begin VB.CommandButton cmdM100 
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
         Height          =   1695
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txtMath 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   4440
         TabIndex        =   1
         Text            =   "Math"
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M100 As String, M200 As String, M300 As String, M400 As String
'Jeopardy.(Jeopardy.vbp)
'Form name: J3; Form caption: Jeopardy
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: This is the Mathematics form. It holds four different buttons that bring
'                up four different questions. Since a lot of the buttons require the same
'                code, we will only explain it once and then if something is new, we will
'                explain it then.

Private Sub cmdM100_Click()
M100 = InputBox("What is 4 + 3?", "M100")           '   This is a a button that will ask a user a
    If M100 = 7 Then                                '   question in the form of an input box.
        MsgBox "Correct, Nice Work!", , "Answer"    '   After answering, a message box will
        Sum = Sum + 100                             '   appear alerting the user whether or not they
    Else                                            '   got the question correct or not.
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 100
    End If
    cmdM100.Visible = False
    If cmdM100.Visible = False And cmdM200.Visible = False And cmdM300.Visible = False And cmdM400.Visible = False Then
        frmMath.Hide                                '   This last part of the button is what makes
        frmTopics.Show                              '   the button dissapear. Once it is answered,
    End If                                          '   it will go away. If all the buttons have been
                                                    '   pressed, it then tells the game to go back to
End Sub                                             '   the main topics form.

Private Sub cmdM200_Click()
M200 = InputBox("What is 12 divided by 3?", "M200")
    If M200 = 4 Then
        MsgBox "Correct, Great Job!", , "Answer"
        Sum = Sum + 200
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 200
    End If
    cmdM200.Visible = False
    If cmdM100.Visible = False And cmdM200.Visible = False And cmdM300.Visible = False And cmdM400.Visible = False Then
        frmMath.Hide
        frmTopics.Show
    End If
End Sub

Private Sub cmdM300_Click()
M300 = InputBox("If you have 6 frogs and find twice as many, how many frogs do you have?", "M300")
    If M300 = 18 Then
        MsgBox "Correct, Terrific!", , "Answer"
        Sum = Sum + 300
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 300
    End If
    cmdM300.Visible = False
    If cmdM100.Visible = False And cmdM200.Visible = False And cmdM300.Visible = False And cmdM400.Visible = False Then
        frmMath.Hide
        frmTopics.Show
    End If
End Sub

Private Sub cmdM400_Click()
M400 = InputBox("If you have 3 bags full of 5 items and take away 2 items, how many total items do you have left?", "M400")
    If M400 = 13 Then
        MsgBox "Correct, You're Amazing!", , "Answer"
        Sum = Sum + 400
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 400
    End If
    cmdM400.Visible = False
    If cmdM100.Visible = False And cmdM200.Visible = False And cmdM300.Visible = False And cmdM400.Visible = False Then
        frmMath.Hide
        frmTopics.Show
    End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturnJ2_Click()
frmTopics.Show
frmMath.Hide
End Sub

Private Sub cmdScore_Click()
    MsgBox "Your score is " & Sum & ". Nice work!"
End Sub

