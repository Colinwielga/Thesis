VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Menu"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   6600
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Play Video Poker!"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End     'Ends Program.
End Sub

Private Sub cmdPlay_Click()
Dim Pos As Integer      'Declares Variable.

frmMain.Hide    'Hides the main screen.
frmMoney.Show    'Opens the monetary denomination screen.

Open App.Path & "/Cards.txt" For Input As #1    'Opens file and enters the 52 cards of the deck into an array.
    For Pos = 1 To 52
        Input #1, Cards(Pos)
    Next Pos
Close #1    'Closes File.


End Sub

