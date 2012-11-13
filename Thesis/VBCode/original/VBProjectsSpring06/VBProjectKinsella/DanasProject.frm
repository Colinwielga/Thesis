VERSION 5.00
Begin VB.Form frmInput 
   BackColor       =   &H80000012&
   Caption         =   "Dana's Music Sorter"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   FillColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   Picture         =   "DanasProject.frx":0000
   ScaleHeight     =   5130
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      Picture         =   "DanasProject.frx":0F1E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New CD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2160
      Picture         =   "DanasProject.frx":1CF0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdmusic 
      Caption         =   "No New CD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2160
      Picture         =   "DanasProject.frx":2AC2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Image Image14 
      Height          =   1680
      Left            =   5040
      Picture         =   "DanasProject.frx":3894
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Image Image13 
      Height          =   1635
      Left            =   3360
      Picture         =   "DanasProject.frx":C856
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Image Image10 
      Height          =   1635
      Left            =   0
      Picture         =   "DanasProject.frx":15440
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Image Image12 
      Height          =   1635
      Left            =   4680
      Picture         =   "DanasProject.frx":1E02A
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Image Image11 
      Height          =   1635
      Left            =   1680
      Picture         =   "DanasProject.frx":26C14
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Image Image9 
      Height          =   1680
      Left            =   1680
      Picture         =   "DanasProject.frx":2F7FE
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Image Image8 
      Height          =   1680
      Left            =   3120
      Picture         =   "DanasProject.frx":387C0
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Image Image7 
      Height          =   1680
      Left            =   0
      Picture         =   "DanasProject.frx":41782
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Image Image6 
      Height          =   1680
      Left            =   4680
      Picture         =   "DanasProject.frx":4A744
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image5 
      Height          =   1635
      Left            =   0
      Picture         =   "DanasProject.frx":53706
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image4 
      Height          =   1635
      Left            =   0
      Picture         =   "DanasProject.frx":5C2F0
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   1635
      Left            =   3120
      Picture         =   "DanasProject.frx":64EDA
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image2 
      Height          =   1680
      Left            =   1560
      Picture         =   "DanasProject.frx":6DAC4
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   0
      Picture         =   "DanasProject.frx":76A86
      Top             =   0
      Width           =   1635
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This button allows the user to exit the program.
Private Sub cmdexit_Click()
    End
End Sub
'This button allows the user to switch from the initial input form
'to the subsequent searching form.
Private Sub cmdmusic_Click()
    frmsearch.Visible = True
    frmInput.Visible = False
End Sub
'This button allows the user to input new information to the text
'file at the beggining of each use.
Private Sub cmdnew_Click()
    Dim NumCDs, NumSongs As Integer
    Dim I, J As Integer
    Open App.Path & "\CDs.txt" For Append As #1
    NumCDs = InputBox("Enter The Number Of New CDs", "New Cds")
    For I = 1 To NumCDs
        GroupArray(I + CDCounter) = InputBox("Enter The Name Of The Group", "Group Name")
        CDArray(I + CDCounter) = InputBox("Enter The Name Of The CD", "CDs")
        NumSongs = InputBox("Enter The Number Of Songs On CD", "Number Of Songs")
        GenreArray(I + CDCounter) = InputBox("Enter The Name Of The Genre", "Genre")
        Write #1, GroupArray(I + CDCounter), CDArray(I + CDCounter), GenreArray(I + CDCounter), NumSongs
        CDCounter = CDCounter + 1
        For J = 1 To NumSongs
            CounterArray(SGCounter + J) = CDCounter
            SongsArray(SGCounter + J) = InputBox("Enter The Name Of The Song", "Songs")
            MoodArray(SGCounter + J) = InputBox("Enter The Mood Of The Song", "Song Mood")
            Write #1, SongsArray(SGCounter + J), MoodArray(SGCounter + J), CounterArray(SGCounter + J)
        Next J
        'Counter = Counter + NumSongs
        'AlbumNum = AlbumNum + 1
    Next I
    Close #1
End Sub
'This button loads all of the arrays to the Form for each use.
Private Sub Form_Load()
    Dim Pos, Pos2 As Integer
    Dim NumberSongs, I As Integer
    Dim Junk As String
    Open App.Path & "\CDs.txt" For Input As #1
    Pos = 0
    Pos2 = 0
    Do While Not EOF(1)
        Pos = Pos + 1
        Input #1, GroupArray(Pos), CDArray(Pos), GenreArray(Pos), NumberSongs
        For I = 1 To Val(NumberSongs)
            Pos2 = Pos2 + 1
            Input #1, SongsArray(Pos2), MoodArray(Pos2), AlbumArray(Pos2)
        Next I
    Loop
    Close #1
    CDCounter = Pos
    SGCounter = Pos2
End Sub

