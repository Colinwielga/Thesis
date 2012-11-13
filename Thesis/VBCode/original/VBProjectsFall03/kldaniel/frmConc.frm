VERSION 5.00
Begin VB.Form frmConc 
   BackColor       =   &H00FF80FF&
   Caption         =   "Bible Concordance"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   ForeColor       =   &H00800080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtWord 
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   2775
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00FFC0FF&
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
      Left            =   240
      ScaleHeight     =   3195
      ScaleWidth      =   6315
      TabIndex        =   4
      Top             =   3840
      Width           =   6375
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   3975
   End
   Begin VB.CommandButton cmdLook 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Verse Search"
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
      Left            =   6960
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   3975
   End
   Begin VB.PictureBox pbxPic 
      Height          =   3735
      Left            =   7200
      Picture         =   "frmConc.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "Where is this word found?"
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
      Left            =   840
      TabIndex        =   6
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label lblConc 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "Bible Concordance"
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
      TabIndex        =   5
      Top             =   120
      Width           =   10815
   End
End
Attribute VB_Name = "frmConc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmConc.Hide
frmWelcome.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLook_Click()
frmConc.Hide
frmVerseSearch.Show
End Sub

Private Sub cmdSearch_Click()
Dim strTemp, strWord, strFind, strSame As String
Dim i, j, l, m As Integer
Dim k As Single
Dim End_Of_Verse, Found As Boolean
pbxResults.Cls
End_Of_Verse = False
strWord = txtWord.Text
i = 1
l = Len(strWord)
For i = 1 To 32
    strTemp = strVerse(1, i)
    'pbxResults.Print strTemp
    k = Len(strTemp)
    Do Until End_Of_Verse = True
        If k <= 0 Then
            End_Of_Verse = True
        End If
        If k > 0 Then
            j = InStr(strTemp, " ")
            strFind = Left(strTemp, j - 1)
            'pbxResults.Print strFind
            strSame = Left(strFind, l)
            If strWord = strSame Then
                pbxResults.Print strBook; strChapter(1); ":"; i
                Found = True
                End_Of_Verse = True
            End If
            strTemp = Right(strTemp, k - j)
            'pbxResults.Print strTemp
            k = k - j
        End If
    Loop
    End_Of_Verse = False
Next i
If Found = False Then
    MsgBox "Sorry, the word is not found", , "Not Found"
Else
    pbxPic = LoadPicture(strPath & "Pictures\scroll3.bmp")
End If
End Sub

