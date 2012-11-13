VERSION 5.00
Begin VB.Form MainPage 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   BeginProperty Font 
      Name            =   "Myriad Pro Light"
      Size            =   36
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "MainPage.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00008000&
      Caption         =   "Trivia Question!"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro B"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "M    I   N N E  S    O T  A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   8055
      Left            =   9000
      TabIndex        =   3
      Top             =   0
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      Caption         =   "Project by Tony Blum and Danielle Johnson  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblMinnesota 
      BackColor       =   &H000040C0&
      Caption         =   "How well ya know Minnesooota??....we will test you to see!.....                                                  ya you betcha!"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   9615
   End
End
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Minnesoooota
'Form Name: MainPage
'Author: Danielle Johnson and Tony Blum
'Date Written: March 26th 2008
'Our purpose for the project is to test the user on how well they know Minnesota and allow the user to explore and experience different tidbidts about Minnesotan culture.
'This is the opening page and the only way for the user to get any futher along is the answer a simple MN trivia question
Option Explicit 'declares all variables

Private Sub cmdStart_Click()
Dim Question1 As String 'question1 is going to be the user's word answer

Question1 = InputBox("To start off easy, what is the official state bird of Minnesota?", "Trivia Question")  'input box is where user types in their answer

If Question1 = "Loon" Or Question1 = "loon" Then
    MsgBox "Why yes, it is the Loon smarty pants!Let's find out some more about our wonderful state!", , "Answer" 'if user types in either those two options then this message box will pop up
    Else
    MsgBox "No, ya silly its the Loon! Uh oh, someone needs to brush off on some MN common knowledge...", , "Answer" ' if user types in anything other than the two options then this message will pop up
    End If

MainPage.Hide 'after one of the messages pop up, then the user moves on to the next form
Minnesota.Show

End Sub

