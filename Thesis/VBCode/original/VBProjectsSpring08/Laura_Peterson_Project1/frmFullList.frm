VERSION 5.00
Begin VB.Form frmFullList 
   Caption         =   "Full List"
   ClientHeight    =   9765
   ClientLeft      =   1830
   ClientTop       =   840
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   Picture         =   "frmFullList.frx":0000
   ScaleHeight     =   9765
   ScaleWidth      =   11085
   Begin VB.CommandButton cmdYearSort 
      BackColor       =   &H000080FF&
      Caption         =   "List All Films in Order of Their Release"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H000080FF&
      Caption         =   "List All Films in Alphabetical Order"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000080FF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H000080FF&
      Caption         =   "First! List All 50 Films in Laura's Gallery"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   9735
      Left            =   3000
      ScaleHeight     =   9675
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmFullList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option explicit
'Laura's Movie Gallery
'frmFullList
'Laura Peterson
'3/27/2008
'this form will print the entire movie list and arrange them alphabetically and by year
Dim CTR As Integer, Film(1 To 50) As String, Year(1 To 50) As Integer
Dim Pass As Integer, TempFilm As String, TempYear As Integer, S As Integer, Pos As Integer
Dim V As Integer
'This will sort the Films Alphabetically
Private Sub cmdAlpha_Click()
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Film(Pos) > Film(Pos + 1) Then
            TempFilm = Film(Pos)
            Film(Pos) = Film(Pos + 1)
            Film(Pos + 1) = TempFilm
            TempYear = Year(Pos)
            Year(Pos) = Year(Pos + 1)
            Year(Pos + 1) = TempYear
        End If
    Next Pos
Next Pass
picResults.Cls
'this will print the results while keeping the film with it's corresponding year
For S = 1 To CTR
    picResults.Print Film(S); Tab(60); Year(S)
Next S
End Sub

Private Sub cmdList_Click()
'This will Load the file of all the films
Open App.Path & "\full list.txt" For Input As #9
Dim K As Integer
CTR = 0
Do While Not EOF(9)
    CTR = CTR + 1
    Input #9, Film(CTR), Year(CTR)
Loop
For K = 1 To CTR
    picResults.Print Film(K);
    picResults.Print Tab(60); Year(K)
Next K


End Sub
'this will hide the full list form and show the genres form
Private Sub cmdReturn_Click()
frmGenres.Show
frmFullList.Hide
End Sub
'this will sort the array by year
Private Sub cmdYearSort_Click()
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Year(Pos) > Year(Pos + 1) Then
            TempYear = Year(Pos)
            Year(Pos) = Year(Pos + 1)
            Year(Pos + 1) = TempYear
            TempFilm = Film(Pos)
            Film(Pos) = Film(Pos + 1)
            Film(Pos + 1) = TempFilm
        End If
    Next Pos
Next Pass
picResults.Cls
For V = 1 To CTR
    picResults.Print Film(V); Tab(60); Year(V)
Next V
End Sub
