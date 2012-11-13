VERSION 5.00
Begin VB.Form frmPS3 
   BackColor       =   &H00000000&
   Caption         =   "Sony Playstation 3"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3735
      Left            =   3360
      ScaleHeight     =   3675
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton cmdRetun 
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
      Left            =   720
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3315
      Left            =   0
      Picture         =   "frmPS3.frx":0000
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmPS3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmPS3
'26 March 2007

Option Explicit
Private Sub cmdRetun_Click()
    frmPS3.Hide             'Hides PS3 form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the Playstation3.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\Playstation3.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Playstation3(Ctr)
            picResults.Print ; Playstation3(Ctr)
            Loop
        Close #1
End Sub

