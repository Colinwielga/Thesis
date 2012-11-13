VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00800000&
   Caption         =   "Intro"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "GO!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "by Lisa Hammer and Kate Bancks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   6360
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   6165
      Left            =   240
      Picture         =   "Intro.frx":0000
      Top             =   360
      Width           =   4200
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your First Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4800
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBegin_Click()                'prjCIRCUSFUN, frmGames by Kate Bancks and Lisa Hammer, March 23, 2006
    N = txtName.Text                        'the purpose of prjCIRCUSFUN is that of an interactive childrens' game with a circus theme.
    frmGames.Show                           'the purpose of frmIntro is to introduce the user to the game and vice versa.
    frmIntro.Hide                           'this button is used to store name and switch forms.
    frmGames.Visible = True
    frmIntro.Visible = False
    
End Sub

Private Sub cmdExit_Click()
    End                                     'this button allows the user to end the program
End Sub
