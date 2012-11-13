VERSION 5.00
Begin VB.Form frmXbox 
   BackColor       =   &H00000000&
   Caption         =   "Microsoft Xbox"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4935
      Left            =   5400
      ScaleHeight     =   4875
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   240
      Width           =   5295
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
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "frmXbox.frx":0000
      Top             =   0
      Width           =   5250
   End
End
Attribute VB_Name = "frmXbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdReturn_Click()
    frmXbox.Hide
    frmConsoleInfo.Show
End Sub
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\XBOX.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, XBOX(Ctr)
            picResults.Print ; XBOX(Ctr)
            Loop
        Close #1
End Sub
