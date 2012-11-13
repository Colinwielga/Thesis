VERSION 5.00
Begin VB.Form frmItalian 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Italian"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmItalian.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmItalian.frx":08CA
   ScaleHeight     =   7845
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click here to see what you need for Pork Roast with Orange"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
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
      ForeColor       =   &H00000040&
      Height          =   3495
      Left            =   4920
      Picture         =   "frmItalian.frx":4B090C
      ScaleHeight     =   3435
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   1920
      Width           =   6615
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H0080C0FF&
      Caption         =   "Show Procedures"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pork Roast with Orange"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   7575
   End
   Begin VB.Image imgItalian 
      Height          =   5760
      Left            =   240
      Picture         =   "frmItalian.frx":5134FE
      Top             =   1440
      Width           =   4575
   End
End
Attribute VB_Name = "frmItalian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim italian(1 To 18) As String
Dim CTR As Integer, I As Integer

Private Sub cmdBack_Click()

frmItalian.Hide
frmCountries.Show

End Sub

Private Sub cmdNext_Click()
groceryfile = "\Recipes\italianR.txt"

'Next Step
frmItalian.Hide
frmGroceryStore.Show

Close #1

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShow_Click()

CTR = 0

Open App.Path & "\italian.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, italian(CTR)
Loop

For I = 1 To CTR
    picResults.Print italian(I)
    
Next I

Close #1

End Sub
