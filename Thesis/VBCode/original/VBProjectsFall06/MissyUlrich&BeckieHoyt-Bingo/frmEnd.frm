VERSION 5.00
Begin VB.Form frmEnd 
   BackColor       =   &H80000007&
   Caption         =   "Thanks For Playing!"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "The End"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   720
      Picture         =   "frmEnd.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   6975
   End
End
Attribute VB_Name = "frmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BingoProject
'End Form
'Missy Ulrich & Beckie Hoyt
'November 3, 2006
'This form simply displays a picture and allows the user to exit the program through an exit command button

Option Explicit

Private Sub cmdQuit_Click()
    End
    'This command button allows the user to end the program
End Sub

Private Sub Form_Load()
    MsgBox "This game was created by two poor college students...do you think we have the money to buy prizes?!?!?", , "Claim Your Prize"
    'This messagebox will appear when the form opens saying the above message to the user
End Sub

