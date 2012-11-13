VERSION 5.00
Begin VB.Form frmGameCube 
   BackColor       =   &H00800000&
   Caption         =   "Nintendo GameCube"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4575
      Left            =   6240
      ScaleHeight     =   4515
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   480
      Width           =   5415
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
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "frmGameCube.frx":0000
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmGameCube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmGameCube
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmGameCube.Hide        'Hides GameCube form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the GameCube.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\GameCube.txt" For Input As #1     'Opens txt document for display
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, GameCube(Ctr)
            picResults.Print ; GameCube(Ctr)
            Loop
        Close #1
End Sub
