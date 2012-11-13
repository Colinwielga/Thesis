VERSION 5.00
Begin VB.Form frmWinner 
   Appearance      =   0  'Flat
   BackColor       =   &H000000C0&
   Caption         =   "Winner!!!"
   ClientHeight    =   8430
   ClientLeft      =   1905
   ClientTop       =   2130
   ClientWidth     =   10470
   DrawMode        =   1  'Blackness
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Palette         =   "frmWinner.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.PictureBox picWinner 
      Height          =   4215
      Left            =   2280
      Picture         =   "frmWinner.frx":158BE
      ScaleHeight     =   4155
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdPrize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to claim your prize!!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   1
      Left            =   7440
      Shape           =   2  'Oval
      Top             =   2280
      Width           =   855
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   0
      Left            =   840
      Shape           =   2  'Oval
      Top             =   2280
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Index           =   1
      Left            =   7320
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Index           =   0
      Left            =   720
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BingoProject
'Winner Form
'Missy Ulrich & Beckie Hoyt
'November 3, 2006
'This form displays a picture to the user to congratulate them for winning the game of bingo
'The form also allows the user to exit the program through a command button
'The user also has the option of using a command button to switch to the End form to claim their prize

Option Explicit

Private Sub cmdExit_Click()
    End
    'This allows the user to exit out of the program
End Sub

Private Sub cmdPrize_Click()
    frmWinner.Hide
    frmEnd.Show
    'This allows the user to switch from the Winner form to the End form to eventually end the program
End Sub


