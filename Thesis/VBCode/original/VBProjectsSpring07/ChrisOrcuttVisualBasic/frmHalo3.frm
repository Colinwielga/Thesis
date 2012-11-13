VERSION 5.00
Begin VB.Form frmHalo3 
   BackColor       =   &H00000000&
   Caption         =   "Halo 3"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   2295
      Left            =   5280
      ScaleHeight     =   2235
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
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
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "frmHalo3.frx":0000
      Top             =   0
      Width           =   6480
   End
End
Attribute VB_Name = "frmHalo3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmHalo3
'26 March 2007

Option Explicit
Private Sub cmdBack_Click()
    frmHalo3.Hide           'Hides Halo3 form
    frmIndustryNews.Show    'Shows IndustryNews form
End Sub
'This command opens the Halo3.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\Halo3.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Halo3(Ctr)
            picResults.Print ; Halo3(Ctr)
            Loop
        Close #1
End Sub
