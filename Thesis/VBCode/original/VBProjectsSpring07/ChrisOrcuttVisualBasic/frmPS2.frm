VERSION 5.00
Begin VB.Form frmPS2 
   BackColor       =   &H00000000&
   Caption         =   "Sony Playstation 2"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4575
      Left            =   5520
      ScaleHeight     =   4515
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
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   5010
      Left            =   0
      Picture         =   "frmPS2.frx":0000
      Top             =   0
      Width           =   5250
   End
End
Attribute VB_Name = "frmPS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmPS2
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmPS2.Hide             'Hides PS2 form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the Playstation2.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
     Dim Ctr As Integer
        Open App.Path & "\Playstation2.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Playstation2(Ctr)
            picResults.Print ; Playstation2(Ctr)
            Loop
        Close #1
End Sub
