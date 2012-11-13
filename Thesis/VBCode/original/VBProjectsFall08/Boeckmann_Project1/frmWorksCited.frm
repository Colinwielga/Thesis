VERSION 5.00
Begin VB.Form frmWorksCited 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Show Works Cited"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   6495
      Left            =   360
      ScaleHeight     =   6435
      ScaleWidth      =   12675
      TabIndex        =   0
      Top             =   240
      Width           =   12735
   End
End
Attribute VB_Name = "frmWorksCited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Scrubs Project
'Works Cited (frmWorksCited)
'Ann Boeckmann
'November 3, 2008
'The purpose of this form is to show the sources for the pictures and info I used


Private Sub cmdBack_Click()

frmWorksCited.Hide
frmOptions.Show

End Sub

Private Sub cmdShow_Click()

Dim Sources As String

Open App.Path & "\sources.txt" For Input As #1

Input #1, Sources

Close #1

picResults.Print Sources

End Sub
