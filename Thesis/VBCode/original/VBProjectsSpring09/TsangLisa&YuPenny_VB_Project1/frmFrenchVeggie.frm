VERSION 5.00
Begin VB.Form frmFrenchVeggie 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "French Vegetarian"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13215
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmFrenchVeggie.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6765
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Click here to see what you need for French Scallops Parisienne"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Procedures"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4695
      Left            =   5040
      Picture         =   "frmFrenchVeggie.frx":08CA
      ScaleHeight     =   4635
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   600
      Width           =   7815
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "French Scallops Parisienne"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   7455
   End
   Begin VB.Image imgFrenchV 
      BorderStyle     =   1  'Fixed Single
      Height          =   7680
      Left            =   -120
      Picture         =   "frmFrenchVeggie.frx":78EAC
      Top             =   -240
      Width           =   5115
   End
End
Attribute VB_Name = "frmFrenchVeggie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frenchveggie(1 To 18) As String
Dim CTR As Integer, FV As Integer

Private Sub cmdBack_Click()

frmFrenchVeggie.Hide
frmCountries.Show

End Sub

Private Sub cmdNext_Click()
groceryfile = "\Recipes\frenchVR.txt"

'Next Step
frmFrenchVeggie.Hide
frmGroceryStore.Show

Close #1

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShow_Click()

CTR = 0

Open App.Path & "\frenchV.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, frenchveggie(CTR)
Loop

For FV = 1 To CTR
    picResults.Print frenchveggie(FV)
    
Next FV

Close #1

End Sub

