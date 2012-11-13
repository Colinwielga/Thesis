VERSION 5.00
Begin VB.Form frmLevel2 
   BackColor       =   &H0080FFFF&
   Caption         =   "Level 2 "
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnswers 
      Caption         =   "Check Answers"
      Height          =   735
      Left            =   6000
      TabIndex        =   15
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtSneetch 
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtGrinch 
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtSam 
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtHorton 
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtCat 
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image imgExit 
      Height          =   705
      Left            =   6840
      Picture         =   "frmLevel2.frx":0000
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   360
      Picture         =   "frmLevel2.frx":0505
      Top             =   2640
      Width           =   1650
   End
   Begin VB.Image Image5 
      Height          =   1605
      Left            =   3360
      Picture         =   "frmLevel2.frx":1433
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1785
   End
   Begin VB.Image Image4 
      Height          =   1815
      Left            =   2760
      Picture         =   "frmLevel2.frx":40595
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2085
   End
   Begin VB.Image Image3 
      Height          =   1935
      Left            =   360
      Picture         =   "frmLevel2.frx":484B0
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2040
   End
   Begin VB.Image Image2 
      Height          =   1485
      Left            =   1680
      Picture         =   "frmLevel2.frx":4D7C5
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label lblDirections 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Match the pictures with the Character's Names.  Use ONLY capitol letters!!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   16
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lnlSneetch 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Sneetches"
      BeginProperty Font 
         Name            =   "Mathematica5Mono"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   14
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblGrinch 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "The Grinch"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblSam 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Sam I Am "
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   12
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblHorton 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Horton"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblCat 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Cat in The Hat"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblE 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "E."
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label lblD 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "D."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "C."
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblB 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "B."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblA 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "A."
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmLevel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnswers_Click()
    Dim Cat As String, Horton As String, Sam As String, Grinch As String, Sneetch As String
    Dim A As String, B As String, C As String, D As String, E As String
    Cat = txtCat.Text               'Inputs value in textbox and sets equal to variable
    Horton = txtHorton.Text         'Inputs value in textbox and sets equal to variable
    Sam = txtSam.Text               'Inputs value in textbox and sets equal to variable
    Grinch = txtGrinch.Text         'Inputs value in textbox and sets equal to variable
    Sneetch = txtSneetch.Text       'Inputs value in textbox and sets equal to variable
    
    If Cat = "B" And Horton = "E" And Sam = "D" And Grinch = "A" And Sneetch = "C" Then     'If each textbox is equal to its variable then "you are correct"
        MsgBox YourName & " Congratulations You are Correct!!", , "Hooray"          'Informs player they have passed the level
        frmLevel2.Visible = False       'level 2 disappears
        frmLevel3.Visible = True        'Level 3 appears
    Else
        MsgBox YourName & " Try Again", , "Oops"        'If player is incorrect, they are then to retry
    End If
        
        
End Sub

Private Sub imgExit_Click()
End           'ends program
End Sub

