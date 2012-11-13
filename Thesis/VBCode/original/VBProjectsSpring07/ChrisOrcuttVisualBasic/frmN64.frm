VERSION 5.00
Begin VB.Form frmN64 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nintendo 64"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   5400
      ScaleHeight     =   4035
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   600
      Width           =   6615
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
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   240
      Picture         =   "frmN64.frx":0000
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "frmN64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmN64
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmN64.Hide             'Hides N64 form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the Nintendo64.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\Nintendo64.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Nintendo64(Ctr)
            picResults.Print ; Nintendo64(Ctr)
            Loop
        Close #1
End Sub
