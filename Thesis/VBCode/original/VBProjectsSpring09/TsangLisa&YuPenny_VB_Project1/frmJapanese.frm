VERSION 5.00
Begin VB.Form frmJapanese 
   Caption         =   "Japanese"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmJapanese.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmJapanese.frx":08CA
   ScaleHeight     =   8880
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Click here to see what you need for Japanese Croquette with Vegetables and Meat"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2895
      Left            =   2880
      Picture         =   "frmJapanese.frx":49150C
      ScaleHeight     =   2835
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   5160
      Width           =   5655
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show Procedures"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Japanese Croquette with Vegetables and Meat"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   8295
   End
   Begin VB.Image imgJapanese 
      Height          =   3675
      Left            =   2880
      Picture         =   "frmJapanese.frx":6D154E
      Top             =   1440
      Width           =   5580
   End
End
Attribute VB_Name = "frmJapanese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim japanese(1 To 18) As String
Dim CTR As Integer, JA As Integer

Private Sub cmdBack_Click()

frmJapanese.Hide
frmCountries.Show

End Sub

Private Sub cmdNext_Click()
groceryfile = "\Recipes\japaneseR.txt"

'Next Step
frmJapanese.Hide
frmGroceryStore.Show

Close #1

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShow_Click()

CTR = 0

Open App.Path & "\japanese.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, japanese(CTR)
Loop

For JA = 1 To CTR
    picResults.Print japanese(JA)
    
Next JA

Close #1

End Sub
