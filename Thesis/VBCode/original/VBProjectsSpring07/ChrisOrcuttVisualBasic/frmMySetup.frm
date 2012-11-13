VERSION 5.00
Begin VB.Form frmMySetup 
   BackColor       =   &H00000000&
   Caption         =   "My Gaming Setup"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   5175
      Left            =   5640
      ScaleHeight     =   5115
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Picture         =   "frmMySetup.frx":0000
      Top             =   0
      Width           =   5400
   End
End
Attribute VB_Name = "frmMySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmMySetup
'26 March 2007

Option Explicit
Private Sub cmdBack_Click()
    frmMySetup.Hide     'Hides Setup form
    frmMyGames.Show     'Shows MyGames form
End Sub
Private Sub cmdReturn_Click()
    frmMySetup.Hide     'Hides Setup form
    frmSelectWant.Show  'Shows SelectWant form
End Sub
'This command opens the MySetup.txt file and displays information
'about the gaming setup featured in the picture
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\MySetup.txt" For Input As #1
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, MySetup(Ctr)
            picResults.Print ; MySetup(Ctr)
            Loop
        Close #1
End Sub
