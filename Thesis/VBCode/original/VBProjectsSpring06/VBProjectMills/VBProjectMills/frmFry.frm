VERSION 5.00
Begin VB.Form frmFry 
   Caption         =   "Bryan Mills"
   ClientHeight    =   8625
   ClientLeft      =   4905
   ClientTop       =   1515
   ClientWidth     =   7680
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmFry.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Enjoy Your Catch"
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "frmFry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is a Fishing guide program (final.vbp)
'Ending form (fry.frm)
'Bryan Mills
'March 24, 2006
'This form tells you where to get recipies and ends the program

Option Explicit

Private Sub cmdQuit_Click()
    MsgBox "For recipies mail $1,000,000 to Bryan Mills @ PO BOX 1242 Collegeville, MN,56321"
    End
    'this message box tells you what to do when you want to cook the fish
    'also ends the program
End Sub
