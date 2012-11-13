VERSION 5.00
Begin VB.Form frmVerseSearch 
   BackColor       =   &H0080FF80&
   Caption         =   "Verse Search"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox txtVerse 
      Height          =   375
      Left            =   10440
      TabIndex        =   13
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtChapter 
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Text            =   "1"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtBook 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Text            =   "Romans"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   4560
      ScaleHeight     =   3195
      ScaleWidth      =   6315
      TabIndex        =   7
      Top             =   3840
      Width           =   6375
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   3975
   End
   Begin VB.CommandButton cmdToConc 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bible Concordance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   3975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Back to Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   3975
   End
   Begin VB.PictureBox pbxPic 
      Height          =   3735
      Left            =   360
      Picture         =   "frmVerseSearch.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblVerse 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "Vs:"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblChapter 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "Ch:"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblBook 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "Book:"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Find this verse!"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Label lblVerseSearch 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Verse Search"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "frmVerseSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
frmVerseSearch.Hide
frmWelcome.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdSearch_Click()
Dim i, j, k, l As Integer
Dim strTemp, strTemp2 As String
Dim End_Of_Verse As Boolean
i = txtChapter.Text
j = txtVerse.Text
End_Of_Verse = False
strTemp = strVerse(i, j)
Do Until End_Of_Verse = True
    l = Len(strTemp)
    If l < 55 Then
        pbxResults.Print strTemp
        End_Of_Verse = True
    Else
        strTemp2 = Mid(strTemp, 45, 10)
        k = InStr(strTemp2, " ")
        pbxResults.Print Left(strTemp, 44 + k - 1)
        strTemp = Right(strTemp, l - (44 + k - 1))
    End If
Loop
pbxResults.Print j
Select Case j
    Case 1 To 6
        pbxPic = LoadPicture(strPath & "Pictures\Gospels.bmp")
    Case 7 To 12
        pbxPic = LoadPicture(strPath & "Pictures\History.bmp")
    Case 13 To 18
        pbxPic = LoadPicture(strPath & "Pictures\Joy.bmp")
    Case 19 To 24
        pbxPic = LoadPicture(strPath & "Pictures\Letters.bmp")
    Case 25 To 32
        pbxPic = LoadPicture(strPath & "Pictures\Prophets.bmp")
End Select
End Sub

Private Sub cmdToConc_Click()
frmVerseSearch.Hide
frmConc.Show
End Sub

Private Sub Command1_Click()
pbxResults.Cls
End Sub

