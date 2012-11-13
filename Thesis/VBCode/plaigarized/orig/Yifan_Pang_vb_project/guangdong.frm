VERSION 5.00
Begin VB.Form guangdong 
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   Picture         =   "guangdong.frx":0000
   ScaleHeight     =   8085
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "return"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox txtpassword 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.PictureBox picresults 
      AutoSize        =   -1  'True
      Height          =   4935
      Left            =   4680
      ScaleHeight     =   4875
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton cmdintroduce 
      BackColor       =   &H8000000D&
      Caption         =   "Csntonese cuisine"
      Height          =   1455
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "3:Fried rice"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "2:Sweet and sour pork"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lab1 
      BackColor       =   &H0080FFFF&
      Caption         =   "1:Chinese steamed eggs"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
End
Attribute VB_Name = "guangdong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Chinese food
'Form Name: guangdong
'Author: Yifan Pang
'Date Written: feb 23 2010
'The purpose of this form is a form to introduce guangdong food
Option Explicit
Private Sub cmdintroduce_Click() 'this is a form read a txt file to picture box
Dim guangdong(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\guangdong.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, guangdong(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.ForeColor = RGB(0, 0, 0) 'use RGB to change font color
    picresults.Print guangdong(n)
Next n
End Sub



Private Sub cmdReturn_Click()
    guangdong.Hide
    China.Show
End Sub

Private Sub Form_Load()

End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer) 'use presskey to show the picture
Dim Char As String, n As Integer
   Char = Chr(KeyAscii)
   KeyAscii = Asc(UCase(Char))
   Select Case Char
    Case "1"
picresults.Picture = LoadPicture(App.Path & "\dan.jpg")
    Case "2"
picresults.Picture = LoadPicture(App.Path & "\rou.jpg")
    Case "3"
picresults.Picture = LoadPicture(App.Path & "\chaofan.jpg")
    Case Else
       MsgBox "not picture in this number", , "error"
    End Select
End Sub

