VERSION 5.00
Begin VB.Form frmCelebrities 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Celebrities"
   ClientHeight    =   8535
   ClientLeft      =   2370
   ClientTop       =   1980
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
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
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox txtCelebrities 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   1320
      TabIndex        =   1
      Text            =   "Celebrities"
      Top             =   360
      Width           =   5535
   End
   Begin VB.PictureBox picJ6 
      BackColor       =   &H00FFFFFF&
      Height          =   8895
      Left            =   0
      Picture         =   "frmJ6.frx":0000
      ScaleHeight     =   8835
      ScaleWidth      =   12675
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      Begin VB.CommandButton cmdScore 
         BackColor       =   &H00FFC0C0&
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6960
         Width           =   2415
      End
      Begin VB.CommandButton cmdTopics 
         BackColor       =   &H00FFC0C0&
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
         Height          =   975
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6960
         Width           =   2175
      End
      Begin VB.CommandButton cmdC400 
         BackColor       =   &H00FFC0C0&
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
         Height          =   1335
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CommandButton cmdC300 
         BackColor       =   &H00FFC0C0&
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
         Height          =   1335
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CommandButton cmdC200 
         BackColor       =   &H00FFC0C0&
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
         Height          =   1335
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2280
         Width           =   2775
      End
      Begin VB.CommandButton cmdC100 
         BackColor       =   &H00FFC0C0&
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
         Height          =   1335
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2280
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmCelebrities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim C100 As String, C200 As String, C300 As String, C400 As String
'Jeopardy.(Jeopardy.vbp)
'Form name: Celebrities; Form caption: Jeopardy
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: This is the Celebrities form. Like before, it displays questions pertaining
'                to celebrities.
    
        

Private Sub cmdC100_Click()
C100 = InputBox("What singer was Jessica Simpson married to? Please answer in all lower case letters", "C100")
    If C100 = "nick lachey" Then
        MsgBox "You are Correct!  Great Answer!", , "Answer"
        Sum = Sum + 100
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 100
    End If
    cmdC100.Visible = False
    If cmdC100.Visible = False And cmdC200.Visible = False And cmdC300.Visible = False And cmdC400.Visible = False Then
        frmCelebrities.Hide
        frmTopics.Show
    End If
End Sub

Private Sub cmdC200_Click()
C200 = InputBox("Name the first name of the youngest member of N'Sync. Please answer in all lower case letters", "C200")
    If C200 = "justin" Then
        MsgBox "You are Correct!  Way to Go!", , "Answer"
        Sum = Sum + 200
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 200
    End If
    cmdC200.Visible = False
    If cmdC100.Visible = False And cmdC200.Visible = False And cmdC300.Visible = False And cmdC400.Visible = False Then
        frmCelebrities.Hide
        frmTopics.Show
    End If
End Sub

Private Sub cmdC300_Click()
C300 = InputBox("What actor played Tony Montana in Scarface? Please answer in all lower case letters", "C300")
    If C300 = "al pacino" Then
        MsgBox "You are Correct!  Nice Work!", , "Answer"
        Sum = Sum + 300
    Else
        MsgBox "Wrong Answer.", , "Wrong Answer"
        Sum = Sum - 300
    End If
    cmdC300.Visible = False
    If cmdC100.Visible = False And cmdC200.Visible = False And cmdC300.Visible = False And cmdC400.Visible = False Then
        frmCelebrities.Hide
        frmTopics.Show
    End If
End Sub

Private Sub cmdC400_Click()
frmCelebrities.Hide             'This button takes the user to another form for the question.
frmCeleb400.Show                'It hides this form and shows the 400 opint question form.
    cmdC400.Visible = False     'After the question is answered, it hides that button.
    If cmdC100.Visible = False And cmdC200.Visible = False And cmdC300.Visible = False And cmdC400.Visible = False Then
        frmCelebrities.Hide
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
frmCelebrities.Hide
frmTopics.Show
End Sub
