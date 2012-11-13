VERSION 5.00
Begin VB.Form frmChoose_Test 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test Prep"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWelcome 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Take Me Back"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "let me browse the works"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "Quiz Me"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmChoose_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'gives options for the user to either review art works or take short quiz thereon

Private Sub cmdBrowse_Click() 'takes user to first browsing form
    frmChoose_Test.Visible = False
    frmBrowse.Visible = True
End Sub

Private Sub cmdQuit_Click() 'opens first form in case the user needs to change his/her name or quit program
    frmChoose_Test.Visible = False
    frmArtHistoryOpen.Visible = True
End Sub


Private Sub cmdQuiz_Click() 'takes user to quiz form
    frmChoose_Test.Visible = False
    frmQuiz.Visible = True
End Sub

Private Sub Form_Activate() 'greets user by name and reads text file on art works
    picWelcome.Cls
   picWelcome.Print Tab(14); "Welcome "; Usr_Name; " ,"
   picWelcome.Print Tab(8); "please choose an option below."
   
   Dim Pos As Integer

Open App.Path & "\Project.txt" For Input As #1

Do Until EOF(1)
    Pos = Pos + 1
    Input #1, workdate(Pos), artists(Pos), titles(Pos), extrainfos(Pos), extrainfos2(Pos)
Loop

Close #1
End Sub



