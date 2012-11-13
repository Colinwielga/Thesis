VERSION 5.00
Begin VB.Form frmCountry1 
   BackColor       =   &H00004080&
   Caption         =   "Country"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Sort list alphabetically."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoadPrint 
      Caption         =   "How many different types of country music styles are there?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000040C0&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   5955
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "Play the Country Concert Wardrobe Quiz"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   6000
      Width           =   975
   End
End
Attribute VB_Name = "frmCountry1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Music VB Project by Cassiann Procenko
'Form Name is frmCountry1
'Date written 10/16/2009
'Purpose of this form is to read the list of country genres in the text document and sort them alphabetically.

Private Sub cmdAlpha_Click()

'clear picture box
picResults.Cls

'define variables
Dim genreList(1 To 50) As String
Dim Tempgenre As String
Dim Pass As Integer, Pos As Integer
Dim CTR As Integer, K As Integer

'open the file to be read and made into arrays
Open App.Path & "\Country.txt" For Input As #1

'create arrays
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, genreList(CTR)
Loop

'sort file
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If genreList(Pos) > genreList(Pos + 1) Then
            Tempgenre = genreList(Pos)
            genreList(Pos) = genreList(Pos + 1)
            genreList(Pos + 1) = Tempgenre
         End If
    Next Pos
Next Pass

'print header information
picResults.Print
picResults.Print "Genres"
picResults.Print "************"

'print results

For K = 1 To CTR
    picResults.Print genreList(K)
Next K

Close #1
End Sub

Private Sub cmdClear_Click()
'clear picture box
picResults.Cls
End Sub

Private Sub cmdLoadPrint_Click()

'clear picture box
picResults.Cls

'define variables
Dim genreList(1 To 50) As String
Dim CTR As Integer

'print header information
picResults.Print
picResults.Print "Genres"
picResults.Print "************"

'open the file to be read and made into arrays
Open App.Path & "\Country.txt" For Input As #1

'create arrays
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, genreList(CTR)
    picResults.Print genreList(CTR)
Loop

Close #1
End Sub

Private Sub cmdQuit_Click()
'show and hide forms
    frmLeave.Show
    frmCountry1.Hide
End Sub

Private Sub cmdQuiz_Click()
'show and hide forms
    frmCountryQuiz.Show
    frmCountry1.Hide
End Sub

Private Sub cmdReturn_Click()
'show and hide forms
    frmMusicTypes.Show
    frmCountry1.Hide
End Sub
