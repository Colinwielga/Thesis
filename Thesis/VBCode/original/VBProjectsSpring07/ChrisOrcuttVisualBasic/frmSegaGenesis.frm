VERSION 5.00
Begin VB.Form frmSegaGenesis 
   BackColor       =   &H00000000&
   Caption         =   "Sega Genesis"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicResults 
      Height          =   2775
      Left            =   1800
      ScaleHeight     =   2715
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   3120
      Width           =   7215
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
      Left            =   4560
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   960
      Picture         =   "frmSegaGenesis.frx":0000
      Top             =   120
      Width           =   9000
   End
End
Attribute VB_Name = "frmSegaGenesis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmSegaGenesis
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmSegaGenesis.Hide     'Hides SegaGenesis form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the SegaGenesis.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\SegaGenesis.txt" For Input As #1
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, SegaGenesis(Ctr)
            picResults.Print ; SegaGenesis(Ctr)
            Loop
        Close #1
End Sub
