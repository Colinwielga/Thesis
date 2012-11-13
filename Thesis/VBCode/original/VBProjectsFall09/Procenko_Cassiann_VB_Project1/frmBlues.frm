VERSION 5.00
Begin VB.Form frmBlues 
   BackColor       =   &H00400000&
   Caption         =   "Blues"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortartist 
      Caption         =   "Sort the songs by artist."
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
      Left            =   4800
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdSorttitle 
      Caption         =   "Sort the songs by title."
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
      Left            =   4800
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdReadandPrint 
      Caption         =   "See list of famous blues songs."
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
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF0000&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6315
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   240
      Width           =   4335
   End
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
      Height          =   735
      Left            =   4800
      TabIndex        =   2
      Top             =   4200
      Width           =   1815
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
      Height          =   735
      Left            =   4800
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
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
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   5880
      Width           =   1815
   End
End
Attribute VB_Name = "frmBlues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Music VB Project by Cassiann Procenko
'Form Name is frmBlues
'Date written 10/16/2009
'Purpose of this form is to read the list of blues songs from the text document, and to sort songs by artist and title

Private Sub cmdClear_Click()
'clear picture box
picResults.Cls
End Sub

Private Sub cmdQuit_Click()
'hide and show forms
    frmLeave.Show
    frmBlues.Hide
End Sub

Private Sub cmdReadandPrint_Click()
'clear picture box
picResults.Cls

'declares variables
Dim songtitleList(1 To 50) As String, artistList(1 To 50) As String
Dim CTR As Integer

'print header information
picResults.Print
picResults.Print "Song Titles"; Tab(30); "Artist"
picResults.Print "*************************************************************************************************"

'open the file to be read and made into arrays
Open App.Path & "\Blues.txt" For Input As #1

'loop to create arrays
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, songtitleList(CTR), artistList(CTR)
    picResults.Print songtitleList(CTR); Tab(30); artistList(CTR)
Loop

Close #1
End Sub

Private Sub cmdReturn_Click()
'hide and show forms
    frmMusicTypes.Show
    frmBlues.Hide
End Sub

Private Sub cmdSortartist_Click()

'clear picture box
picResults.Cls

'define variables
Dim songtitleList(1 To 50) As String, artistList(1 To 50) As String
Dim CTR As Integer, Pos As Integer, Pass As Integer
Dim Tempsongtitle As String, Tempartist As String
Dim K As Integer

'open the file to be read and made into arrays
Open App.Path & "\Blues.txt" For Input As #1

'start reading file
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, songtitleList(CTR), artistList(CTR)
Loop

'sort file
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If artistList(Pos) > artistList(Pos + 1) Then
            Tempartist = artistList(Pos)
            artistList(Pos) = artistList(Pos + 1)
            artistList(Pos + 1) = Tempartist
                Tempsongtitle = songtitleList(Pos)
                songtitleList(Pos) = songtitleList(Pos + 1)
                songtitleList(Pos + 1) = Tempsongtitle
        End If
    Next Pos
Next Pass

'print header
picResults.Print
picResults.Print "Song Titles"; Tab(30); "Artists"
picResults.Print "*******************************************************************************************************"

'print results
For K = 1 To CTR
    picResults.Print songtitleList(K); Tab(30); artistList(K)
Next K

Close #1
End Sub

Private Sub cmdSorttitle_Click()

'clear picture box
picResults.Cls

'define variables
Dim songtitleList(1 To 50) As String, artistList(1 To 50) As String
Dim CTR As Integer, Pos As Integer, Pass As Integer
Dim Tempsongtitle As String, Tempartist As String
Dim K As Integer

'open the file to be read and made into arrays
Open App.Path & "\Blues.txt" For Input As #1

'create arrays
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, songtitleList(CTR), artistList(CTR)
Loop

'sort file
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If songtitleList(Pos) > songtitleList(Pos + 1) Then
            Tempsongtitle = songtitleList(Pos)
            songtitleList(Pos) = songtitleList(Pos + 1)
            songtitleList(Pos + 1) = Tempsongtitle
                Tempartist = artistList(Pos)
                artistList(Pos) = artistList(Pos + 1)
                artistList(Pos + 1) = Tempartist
        End If
    Next Pos
Next Pass

'print header
picResults.Print
picResults.Print "Song Titles"; Tab(30); "Artists"
picResults.Print "*****************************************************************************************"

'print results
For K = 1 To CTR
    picResults.Print songtitleList(K); Tab(30); artistList(K)
Next K

Close #1
End Sub
