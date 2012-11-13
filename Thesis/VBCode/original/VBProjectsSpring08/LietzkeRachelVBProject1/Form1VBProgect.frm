VERSION 5.00
Begin VB.Form frmIreland3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ireland 3"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   Picture         =   "Form1VBProgect.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   11475
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H80000012&
      Caption         =   "Show Picture"
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdGoHome 
      BackColor       =   &H00000000&
      Caption         =   "Go Home"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   7320
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   2520
      ScaleHeight     =   6915
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label lblFileName 
      Caption         =   "File Name will be display here."
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label lblPic7 
      BackColor       =   &H00000000&
      Caption         =   "      7. More              Ireland"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lblPic5 
      BackColor       =   &H00000000&
      Caption         =   "     5. More                Ireland"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblPic6 
      BackColor       =   &H00000000&
      Caption         =   "        6.  More                  Ireland"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblPic4 
      BackColor       =   &H00000000&
      Caption         =   "      4. Southern               Ireland"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblPic3 
      BackColor       =   &H00000000&
      Caption         =   "    3.  Northern         Ireland"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblPic2 
      BackColor       =   &H00000000&
      Caption         =   "         2. Dublin"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblPic1 
      BackColor       =   &H00000000&
      Caption         =   "     1. Galway"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblDirections 
      BackColor       =   &H00000000&
      Caption         =   "Write the Number of the picture you want to see in the Text Box Below!"
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmIreland3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Information About Ireland
'Form Name: Ireland3
'Author: Rachel Lietzke
'Date Written: March 27, 2008
'Objective: To Show what Ireland looks like through Pictrues

Private Sub cmdGoHome_Click()
frmIreland3.Hide
frmIreland1.Show
End Sub
Private Sub cmdShow_Click()
Dim pictureNumber As Integer
Dim CTR As Integer
Dim Picture(1 To 15) As String

CTR = 0
pictureNumber = txtNumber.Text

Open App.Path & "\Pictures.txt" For Input As #1

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Picture(CTR)
Loop

lblFileName.Caption = Picture(pictureNumber)
lblFileName.Visible = True
picResults.Picture = LoadPicture(App.Path & "\" & Picture(pictureNumber))

Close #1

End Sub

