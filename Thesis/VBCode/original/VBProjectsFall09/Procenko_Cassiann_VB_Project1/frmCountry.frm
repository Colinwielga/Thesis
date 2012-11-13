VERSION 5.00
Begin VB.Form frmCountryQuiz 
   BackColor       =   &H0080C0FF&
   Caption         =   "Country"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnswers 
      BackColor       =   &H0080C0FF&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.OptionButton optBlue 
      BackColor       =   &H0080C0FF&
      Caption         =   "blue jeans"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton optJersey 
      BackColor       =   &H0080C0FF&
      Caption         =   "baseball jersey"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.OptionButton optCow 
      BackColor       =   &H0080C0FF&
      Caption         =   "cowboy hat and boots"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblWear 
      BackColor       =   &H0080C0FF&
      Caption         =   "What do you wear to a country music concert?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmCountryQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Music VB Project by Cassiann Procenko
'Form Name is frmCountryQuiz
'Date written 10/16/2009
'Purpose of this form is to ask the viewer a question and allow him/her to choose the answer.  Also gives feedback to the viewer based on his/her choice.

Private Sub cmdAnswers_Click()
'if then statement to separate answers
If optBlue.Value = True Then
        MsgBox "If you wear your jeans to the concert, you'll fit in with everyone else!"
    ElseIf optJersey.Value = True Then
        MsgBox "Go for it!  Wear your baseball jersey to the game.  I bet people will ask you who your favorite team is and talk to you about different sports."
    ElseIf optCow.Value = True Then
         MsgBox "Sure, you can wear these to a country music concert, although the audience may think you are a performer rather than a spectator!"
    Else
        MsgBox "Please click an option."
End If
End Sub

Private Sub cmdQuit_Click()
'show and hide forms
    frmLeave.Show
    frmCountryQuiz.Hide
End Sub

Private Sub cmdReturn_Click()
'show and hide forms
    frmCountry1.Show
    frmCountryQuiz.Hide
End Sub

