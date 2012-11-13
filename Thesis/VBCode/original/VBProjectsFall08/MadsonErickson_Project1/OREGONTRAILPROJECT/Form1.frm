VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Don't Play Oregon Trail (Because you just don't have the spirit of adventure today.)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      TabIndex        =   1
      Top             =   7080
      Width           =   4095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Oregon Trail! "
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   2040
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is the starting page of Oregon trail. It's pretty simple

Private Sub cmdPlay_Click()

Dim UserName As String

    UserName = InputBox("What's your name partner? If you reckon' to hit the the trail, we better know what name to put on your grave.", "Welcome!") 'Retrieves and stores UserName in module/public
    Form1.Hide 'hides start page from user
    Form2.Show 'shows main page to user
    MsgBox "Welcome to the greatest thing you'll ever do, " & UserName & ".", , "Salutations."
End Sub

    
Private Sub cmdQuit_Click()

    End

End Sub
