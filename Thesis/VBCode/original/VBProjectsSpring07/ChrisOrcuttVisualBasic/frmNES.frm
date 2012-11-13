VERSION 5.00
Begin VB.Form frmNES 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nintendo Entertainment System"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12825
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   12825
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      Height          =   4935
      Left            =   5520
      ScaleHeight     =   4875
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
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
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   0
      Picture         =   "frmNES.frx":0000
      Top             =   0
      Width           =   5250
   End
End
Attribute VB_Name = "frmNES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmNES
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmNES.Hide             'Hides NES form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the NES.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\NES.txt" For Input As #1
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, NES(Ctr)
            picResults.Print ; NES(Ctr)
            Loop
        Close #1
End Sub

