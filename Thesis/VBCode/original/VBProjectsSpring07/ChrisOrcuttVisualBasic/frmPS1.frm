VERSION 5.00
Begin VB.Form frmPS1 
   BackColor       =   &H00000000&
   Caption         =   "Sony Playstation"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   4920
      ScaleHeight     =   4035
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   480
      Width           =   5655
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
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   4200
      Left            =   0
      Picture         =   "frmPS1.frx":0000
      Top             =   0
      Width           =   4230
   End
End
Attribute VB_Name = "frmPS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdReturn_Click()
    frmPS1.Hide             'Hides PS1 form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the Playstation.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\Playstation.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Playstation(Ctr)
            picResults.Print ; Playstation(Ctr)
            Loop
        Close #1
End Sub
