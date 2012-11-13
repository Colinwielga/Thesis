VERSION 5.00
Begin VB.Form frmWii 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nintendo Wii"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4575
      Left            =   6000
      ScaleHeight     =   4515
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   720
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
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   4950
      Left            =   0
      Picture         =   "frmWii.frx":0000
      Top             =   0
      Width           =   5580
   End
End
Attribute VB_Name = "frmWii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmWii
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmWii.Hide             'Hides Wii form
    frmConsoleInfo.Show     'Shows ConsoleInfo form
End Sub
'This command opens the NintendoWii.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\NintendoWii.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, NintendoWii(Ctr)
            picResults.Print ; NintendoWii(Ctr)
            Loop
        Close #1
End Sub
