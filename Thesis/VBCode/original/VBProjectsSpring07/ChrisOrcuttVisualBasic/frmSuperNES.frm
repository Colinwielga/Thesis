VERSION 5.00
Begin VB.Form frmSuperNES 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Super Nintendo Entertainment System"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12720
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   12720
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4575
      Left            =   6600
      ScaleHeight     =   4515
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   480
      Width           =   5895
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
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   0
      Picture         =   "frmSuperNES.frx":0000
      Top             =   0
      Width           =   6510
   End
End
Attribute VB_Name = "frmSuperNES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmSuperNES
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmSuperNES.Hide        'Hides SuperNES form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the SuperNintendo.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\SuperNintendo.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, SuperNintendo(Ctr)
            picResults.Print ; SuperNintendo(Ctr)
            Loop
        Close #1
End Sub
