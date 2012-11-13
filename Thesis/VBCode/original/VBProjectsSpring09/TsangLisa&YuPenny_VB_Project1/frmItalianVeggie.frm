VERSION 5.00
Begin VB.Form frmItalianVeggie 
   BackColor       =   &H00000000&
   Caption         =   "Italian Vegetarian"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmItalianVeggie.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmItalianVeggie.frx":08CA
   ScaleHeight     =   8790
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Click here to see what you need for Spaghetti with Olives and Tomatoes"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   855
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
      ForeColor       =   &H000040C0&
      Height          =   4455
      Left            =   4800
      Picture         =   "frmItalianVeggie.frx":27258C
      ScaleHeight     =   4395
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   1920
      Width           =   5055
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Show Procedures"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Spaghetti with olives and tomatoes"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   9855
   End
   Begin VB.Image imgItalian 
      Height          =   4500
      Left            =   240
      Picture         =   "frmItalianVeggie.frx":4E424E
      Top             =   1920
      Width           =   4500
   End
End
Attribute VB_Name = "frmItalianVeggie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim italaianveggie(1 To 18) As String
Dim CTR As Integer, IV As Integer

Private Sub cmdBack_Click()

frmItalianVeggie.Hide
frmCountries.Show

End Sub

Private Sub cmdNext_Click()
groceryfile = "\Recipes\italianVR.txt"

'Next Step
frmItalianVeggie.Hide
frmGroceryStore.Show

Close #1

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShow_Click()

CTR = 0

Open App.Path & "\italianV.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, italaianveggie(CTR)
Loop

For IV = 1 To CTR
    picResults.Print italaianveggie(IV)
    
Next IV

Close #1

End Sub



