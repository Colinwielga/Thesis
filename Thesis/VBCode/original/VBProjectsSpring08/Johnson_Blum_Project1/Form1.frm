VERSION 5.00
Begin VB.Form MainPage 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
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
Option Explicit
Private Sub cmdStart_Click()
Dim Question1 As String

Question1 = InputBox("To start off easy, what is the official state bird of Minnesota?", "Trivia Question")

If Question1 = "Loon" Or Question1 = "loon" Then
    MsgBox ("Why yes, it is the Loon smarty pants!Let's find out some more about our wonderful state!")
    Else
    MsgBox ("No, ya silly its the Loon! Uh oh, someone needs to brush off on some MN common knowledge...")
    End If

MainPage.Hide
Minnesota.Show

End Sub

