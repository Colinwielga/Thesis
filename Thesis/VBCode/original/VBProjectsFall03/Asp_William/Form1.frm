VERSION 5.00
Begin VB.Form FindTwin 
   BackColor       =   &H8000000D&
   Caption         =   "                     Find Your Favorite Minnesota Twins Hitter's Statistics (Bill Asp)"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12630
   ScaleMode       =   0  'User
   ScaleWidth      =   17685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Clear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   8
      Top             =   6600
      Width           =   1695
   End
   Begin VB.PictureBox PicResults2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   720
      ScaleHeight     =   1995
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton RBI 
      Caption         =   "RBI Team Leaders for 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton HomeRuns 
      Caption         =   "Home Run Team Leaders for 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton AtBatLeaders 
      Caption         =   "Team Leaders In At Bats for 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton FindTwin 
      Caption         =   "Find Your Twins Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      TabIndex        =   1
      Top             =   6600
      Width           =   1575
   End
   Begin VB.PictureBox PicResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   2760
      ScaleHeight     =   5595
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   600
      Width           =   9975
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Project By Bill Asp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4575
      Left            =   2760
      TabIndex        =   9
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Find Your Favorite Minnesota Twins Hitter's Statistics for 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "FindTwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'VB Project(M:\CS130\William Asp)
'FindTwin (Form1.frm)
'William Asp
'October 28, 2003
'This project has the purpose of finding different players for the Minnesota Twins
'professional baseball team and being able to see not only their major statistics,
'but also to see their picture and where they rank amongst their teammates in major
'statistical categories.
Dim Path As String


Private Sub AtBatLeaders_Click() 'sorts players by AtBats in 2003 season.
Path = "N:\CS130\handin\Asp_William\"
Dim Players(1 To 17) As String, Position(1 To 17) As String
Dim AtBats(1 To 17) As Integer, Hits(1 To 17) As Integer
Dim HomeRuns(1 To 17) As Integer, RBI(1 To 17) As Integer, Pics(1 To 17) As String
Dim Pass As Integer, Temp As String, I As Integer, X As Integer
Open Path & "stats.txt" For Input As #1

PicResults.Cls
PicResults2.Cls
PicResults.Print "Player"; Tab(23); "AtBats"
PicResults.Print "--------------------------------------------------------"

For I = 1 To 17
    Input #1, Players(I), Position(I), AtBats(I), Hits(I), HomeRuns(I), RBI(I), Pics(I)
Next I

    For Pass = 1 To 16 'This bubble sort takes atbats and makes that sortable along with player's names.
        For I = 1 To 17 - Pass
            If AtBats(I) < AtBats(I + 1) Then
                Temp = AtBats(I)
                AtBats(I) = AtBats(I + 1)
                AtBats(I + 1) = Temp
                Temp = Players(I)
                Players(I) = Players(I + 1)
                Players(I + 1) = Temp
            End If
        Next I
    Next Pass

    For I = 1 To 17
        PicResults.Print Players(I); Tab(23); AtBats(I) 'prints player's name and atbats
        PicResults2.Picture = LoadPicture(Pics(I)) 'prints leading player's picture
    Next I
Close #1
End Sub

Private Sub Clear_Click()
PicResults.Cls
Set PicResults2.Picture = LoadPicture
Close #1
End Sub

Private Sub FindTwin_Click() 'This allows user to type in favorite player into an input box.
Dim Players(1 To 17) As String, Position(1 To 17) As String
Dim AtBats(1 To 17) As Integer, Hits(1 To 17) As Integer
Dim HomeRuns(1 To 17) As Integer, RBI(1 To 17) As Integer, Pics(1 To 17) As String
Dim N As String, I As Integer, X As Integer
Dim NotFound As Boolean

Open Path & "stats.txt" For Input As #1

PicResults.Cls

For I = 1 To 17
    Input #1, Players(I), Position(I), AtBats(I), Hits(I), HomeRuns(I), RBI(I), Pics(I)
Next I

N = InputBox("Enter your Twins player.", "Players") 'Brings up input box and asks user for name.
I = 1
NotFound = True
Do While NotFound 'This provides a sequential search throughout the array for the proper name.
    If I >= 18 Then
       Exit Do
    Else
        If N = Players(I) Then
        NotFound = False
        X = I
        Exit Do
        End If
        I = I + 1
    End If
    Loop
If NotFound Then 'This tells the user there was an error and what caused it.
        MsgBox "Sorry but this person is not part of the Twins organization", , "Wrong Player"
        MsgBox "Remember to use capital letters for first and last names and to spell the player's full first and last names correctly", , "Remember"
        Close #1
    Else
        PicResults.Print "Player"; Tab(20), "Position", "At Bats", "Hits", "Home Runs", "RBIs"
        PicResults.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PicResults.Print N; Tab(20), Position(X), AtBats(X), Hits(X), HomeRuns(X), RBI(X) ' Prints player and his statistics
        PicResults2.Picture = LoadPicture(Pics(X)) 'Prints player's picture
Close #1
End If
End Sub

Private Sub HomeRuns_Click() 'sorts players by Home Runs in 2003 season
Dim Players(1 To 17) As String, Position(1 To 17) As String
Dim AtBats(1 To 17) As Integer, Hits(1 To 17) As Integer
Dim HomeRuns(1 To 17) As Integer, RBI(1 To 17) As Integer, Pics(1 To 17) As String
Dim Pass As Integer, Temp As String, I As Integer
Open Path & "stats.txt" For Input As #1

PicResults.Cls
PicResults2.Cls
PicResults.Print "Player"; Tab(23); "Home Runs"
PicResults.Print "--------------------------------------------------------"

For I = 1 To 17
    Input #1, Players(I), Position(I), AtBats(I), Hits(I), HomeRuns(I), RBI(I), Pics(I)
Next I

    For Pass = 1 To 16 'This bubble sort takes Home Runs and makes that sortable along with player's names.
        For I = 1 To 17 - Pass
            If HomeRuns(I) < HomeRuns(I + 1) Then
                Temp = HomeRuns(I)
                HomeRuns(I) = HomeRuns(I + 1)
                HomeRuns(I + 1) = Temp
                Temp = Players(I)
                Players(I) = Players(I + 1)
                Players(I + 1) = Temp
            End If
        Next I
    Next Pass

    For I = 1 To 17
        PicResults.Print Players(I); Tab(23); HomeRuns(I) 'Prints player name and homeruns
        PicResults2.Picture = LoadPicture(Pics(I)) 'Prints leading player's picture
    Next I
Close #1
End Sub

Private Sub Quit_Click()
End
End Sub


Private Sub RBI_Click() 'sorts players by RBIs in 2003 season
Dim Players(1 To 17) As String, Position(1 To 17) As String
Dim AtBats(1 To 17) As Integer, Hits(1 To 17) As Integer
Dim HomeRuns(1 To 17) As Integer, RBI(1 To 17) As Integer, Pics(1 To 17) As String
Dim Pass As Integer, Temp As String, I As Integer
Open Path & "stats.txt" For Input As #1

PicResults.Cls
PicResults2.Cls
PicResults.Print "Player"; Tab(23); "RBIs"
PicResults.Print "--------------------------------------------------------"

For I = 1 To 17
    Input #1, Players(I), Position(I), AtBats(I), Hits(I), HomeRuns(I), RBI(I), Pics(I)
Next I

    For Pass = 1 To 16 'This bubble sort takes RBIs and makes that sortable along with player's names
        For I = 1 To 17 - Pass
            If RBI(I) < RBI(I + 1) Then
                Temp = RBI(I)
                RBI(I) = RBI(I + 1)
                RBI(I + 1) = Temp
                Temp = Players(I)
                Players(I) = Players(I + 1)
                Players(I + 1) = Temp
            End If
        Next I
    Next Pass

    For I = 1 To 17
        PicResults.Print Players(I); Tab(23); RBI(I) 'Prints player name and RBIs
        PicResults2.Picture = LoadPicture(Pics(I)) 'Prints leading player's picture
    Next I
Close #1
End Sub
