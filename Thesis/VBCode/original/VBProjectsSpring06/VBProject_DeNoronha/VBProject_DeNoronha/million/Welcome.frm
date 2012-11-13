VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00400000&
   Caption         =   "Welcome"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10860
   LinkTopic       =   "Form2"
   ScaleHeight     =   7275
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdplay 
      Caption         =   "Lets Play"
      BeginProperty Font 
         Name            =   "Edwardian Script ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3240
      TabIndex        =   1
      Top             =   5280
      Width           =   3735
   End
   Begin VB.PictureBox piclogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   2400
      Picture         =   "Welcome.frx":0000
      ScaleHeight     =   3945
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label lbDesigner 
      BackColor       =   &H00400000&
      Caption         =   "Pradeep de Noronha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
End
Attribute VB_Name = "frmwelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Who wants to be a Millionare.(millionare1.vbp)

'Form name: frmWelcome; Form caption: Welcome

'Author: Pradeep de Noronha

'Date written: 15th March, 2006

'Program Objective: The following program is a VB version of the game show
'                   "Who wants to be a Millionare." The program exhibits all
'                   the features that would be seen on the show. It asks the
'                   user a set of fifteen question and if answered right the
'                   user could leave with a Million dollars.

' Form Objective: The frmWelcome form is the opening screen for the game show.
'                 Its ask the user to input his or her name via an InputBox
'                 and then guides them to the next form(frmMillionare) where
'                 the user starts playing the game.

Private Sub cmdplay_Click()
    username = InputBox("Please enter your name.", "User Input")
    frmmillionare.Show
    frmwelcome.Hide
        
End Sub


