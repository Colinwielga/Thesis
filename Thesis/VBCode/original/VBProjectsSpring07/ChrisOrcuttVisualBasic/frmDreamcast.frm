VERSION 5.00
Begin VB.Form frmDreamcast 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sega Dreamcast"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4335
      Left            =   6720
      ScaleHeight     =   4275
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   360
      Width           =   5175
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
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4665
      Left            =   0
      Picture         =   "frmDreamcast.frx":0000
      Top             =   0
      Width           =   6300
   End
End
Attribute VB_Name = "frmDreamcast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmDreamcast
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmDreamcast.Hide       'Hide Dreamcast form
    frmConsoleInfo.Show     'Show ConsoleInfo form
End Sub
'This command opens the SegaDreamcast.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
     Dim Ctr As Integer
        Open App.Path & "\SegaDreamcast.txt" For Input As #1    'Opens txt document for display
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, SegaDreamcast(Ctr)
            picResults.Print ; SegaDreamcast(Ctr)
            Loop
        Close #1
End Sub
