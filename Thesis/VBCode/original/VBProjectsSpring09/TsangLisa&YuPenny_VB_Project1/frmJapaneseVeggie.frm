VERSION 5.00
Begin VB.Form frmJapaneseVeggie 
   Caption         =   "Japanese Vegetarian"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmJapaneseVeggie.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmJapaneseVeggie.frx":08CA
   ScaleHeight     =   7275
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to see what you need for Kombu Dashi Soup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   3720
      Picture         =   "frmJapaneseVeggie.frx":16020C
      ScaleHeight     =   3675
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Procedures"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H8000000E&
      Caption         =   "Kombu Dashi Soup"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   4335
   End
   Begin VB.Image imgJapaneseV 
      Height          =   3735
      Left            =   240
      Picture         =   "frmJapaneseVeggie.frx":2860C6
      Top             =   1800
      Width           =   3270
   End
End
Attribute VB_Name = "frmJapaneseVeggie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim japaneseveggie(1 To 18) As String
Dim CTR As Integer, JAV As Integer

Private Sub cmdBack_Click()

frmJapaneseVeggie.Hide
frmCountries.Show

End Sub

Private Sub cmdNext_Click()
groceryfile = "\Recipes\japaneseVR.txt"

'Next Step
frmJapaneseVeggie.Hide
frmGroceryStore.Show

Close #1

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShow_Click()

CTR = 0

Open App.Path & "\japaneseV.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, japaneseveggie(CTR)
Loop

For JAV = 1 To CTR
    picResults.Print japaneseveggie(JAV)
    
Next JAV

Close #1

End Sub

