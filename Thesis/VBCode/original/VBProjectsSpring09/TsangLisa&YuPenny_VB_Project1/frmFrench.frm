VERSION 5.00
Begin VB.Form frmFrench 
   BackColor       =   &H00FFFFFF&
   Caption         =   "French"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
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
   MouseIcon       =   "frmFrench.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmFrench.frx":08CA
   ScaleHeight     =   7830
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Click here to see what you need for Blue Ribbon Chicken"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Recipe"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
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
      Height          =   2775
      Left            =   3840
      Picture         =   "frmFrench.frx":49150C
      ScaleHeight     =   2715
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Image imgFrench 
      Height          =   3330
      Left            =   3840
      Picture         =   "frmFrench.frx":5BC1A6
      Top             =   4320
      Width           =   5985
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Blue Ribbon Chicken"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmFrench"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim french(1 To 18) As String
Dim CTR As Integer, F As Integer


Private Sub cmdBack_Click()

frmFrench.Hide
frmCountries.Show

End Sub

Private Sub cmdNext_Click()
groceryfile = "\Recipes\frenchR.txt"

'Next Step
frmFrench.Hide
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
    Input #1, french(CTR)
Loop

For F = 1 To CTR
    picResults.Print french(F)
    
Next F

Close #1

End Sub

