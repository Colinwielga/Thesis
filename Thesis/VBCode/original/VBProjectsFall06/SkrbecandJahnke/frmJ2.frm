VERSION 5.00
Begin VB.Form frmTopics 
   Caption         =   "Jeopardy"
   ClientHeight    =   9480
   ClientLeft      =   1470
   ClientTop       =   915
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPlayGame 
      BackColor       =   &H000080FF&
      Height          =   9615
      Left            =   0
      Picture         =   "frmJ2.frx":0000
      ScaleHeight     =   9555
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdGrandTotal 
         BackColor       =   &H00FF8080&
         Caption         =   "Click to View Your Grand Total!"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6960
         Width           =   3615
      End
      Begin VB.CommandButton cmdTotal 
         BackColor       =   &H00FF8080&
         Caption         =   "Click to View Your  Current Score"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5880
         Width           =   2415
      End
      Begin VB.CommandButton cmdSports 
         BackColor       =   &H00FFFF80&
         Caption         =   "Sports"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CommandButton cmdCelebrities 
         BackColor       =   &H00FFFF80&
         Caption         =   "Celebrities"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   2535
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FF8080&
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
         Height          =   615
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8760
         Width           =   1815
      End
      Begin VB.CommandButton History 
         BackColor       =   &H00FFFF80&
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton cmdGeography 
         BackColor       =   &H00FFFF80&
         Caption         =   "Geography"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtTopic 
         BackColor       =   &H000080FF&
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
         Left            =   3000
         TabIndex        =   2
         Text            =   "Pick A Topic"
         Top             =   240
         Width           =   6495
      End
      Begin VB.CommandButton cmdMath 
         BackColor       =   &H00FFFF80&
         Caption         =   "Mathematics"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2040
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmTopics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Jeopardy.(Jeopardy.vbp)
'Form name: Topics; Form caption: Jeopardy
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: This is the form that allows the user to pick between five (5) different
'                categories. These categories were based on the actual Jeopardy show.
'                Users can choose between Mathematics, Geography, History, Celebrities,
'                and Sports. All of the buttons on this form will take them to the form
'                of the category they chose.

Private Sub cmdCelebrities_Click()      ' This button takes you to the Celebrities form.
frmTopics.Hide
frmCelebrities.Show
End Sub

Private Sub cmdGeography_Click()        ' This button takes you to the Geography form.
frmTopics.Hide
frmGeography.Show
End Sub

Private Sub cmdGrandTotal_Click()       ' This button takes you to the GrandTotal form.
frmTopics.Hide
frmGrandTotal.Show
End Sub

Private Sub cmdMath_Click()             ' This button takes you to the Math form.
frmTopics.Hide
frmMath.Show
End Sub

Private Sub cmdQuit_Click()             ' This button quits the program.
End
End Sub

Private Sub cmdSports_Click()           ' This button takes you to the Sports form.
frmTopics.Hide
frmSports.Show
End Sub

Private Sub cmdTotal_Click()            ' This button shows you your current score.
    MsgBox "Your score is " & Sum & ". Nice work!"
End Sub

Private Sub History_Click()             ' This button takes you to the History form.
frmTopics.Hide
frmHistory.Show
End Sub

