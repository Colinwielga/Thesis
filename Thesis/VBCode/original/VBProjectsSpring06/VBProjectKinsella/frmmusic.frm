VERSION 5.00
Begin VB.Form frmsearch 
   BackColor       =   &H8000000D&
   Caption         =   "Music"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   FillColor       =   &H00C00000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "frmmusic.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdback 
      Caption         =   "Back To Previous Page"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmmusic.frx":B8C2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox picOutput 
      AutoRedraw      =   -1  'True
      Height          =   3975
      Left            =   2640
      ScaleHeight     =   3915
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   360
      Width           =   3375
      Begin VB.Image Image17 
         Height          =   15
         Left            =   1200
         Top             =   3960
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdmood 
      Caption         =   "Search For Mood"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmmusic.frx":C320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdsong 
      BackColor       =   &H80000001&
      Caption         =   "Search For Song"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmmusic.frx":CD7E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Picture         =   "frmmusic.frx":D7DC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdgroup 
      Caption         =   "Search For Group"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "frmmusic.frx":E23A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Image Image20 
      Height          =   1425
      Left            =   5760
      Picture         =   "frmmusic.frx":EC98
      Top             =   4320
      Width           =   1905
   End
   Begin VB.Image Image19 
      Height          =   375
      Left            =   5760
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image18 
      Height          =   375
      Left            =   3840
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Image Image16 
      Height          =   1425
      Left            =   6000
      Picture         =   "frmmusic.frx":17B5A
      Top             =   2880
      Width           =   1905
   End
   Begin VB.Image Image15 
      Height          =   1425
      Left            =   6000
      Picture         =   "frmmusic.frx":20A1C
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Image Image14 
      Height          =   1425
      Left            =   5760
      Picture         =   "frmmusic.frx":298DE
      Top             =   0
      Width           =   1905
   End
   Begin VB.Image Image13 
      Height          =   1215
      Left            =   3840
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image Image12 
      Height          =   1425
      Left            =   0
      Picture         =   "frmmusic.frx":327A0
      Top             =   3360
      Width           =   1905
   End
   Begin VB.Image Image11 
      Height          =   1425
      Left            =   0
      Picture         =   "frmmusic.frx":3B662
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Image Image10 
      Height          =   1425
      Left            =   0
      Picture         =   "frmmusic.frx":44524
      Top             =   0
      Width           =   1905
   End
   Begin VB.Image Image9 
      Height          =   1095
      Left            =   1920
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Image Image8 
      Height          =   1425
      Left            =   1920
      Picture         =   "frmmusic.frx":4D3E6
      Top             =   2760
      Width           =   1905
   End
   Begin VB.Image Image7 
      Height          =   1425
      Left            =   1920
      Picture         =   "frmmusic.frx":562A8
      Top             =   3720
      Width           =   1905
   End
   Begin VB.Image Image6 
      Height          =   1425
      Left            =   0
      Picture         =   "frmmusic.frx":5F16A
      Top             =   1920
      Width           =   1905
   End
   Begin VB.Image Image5 
      Height          =   1425
      Left            =   1920
      Picture         =   "frmmusic.frx":6802C
      Top             =   0
      Width           =   1905
   End
   Begin VB.Image Image4 
      Height          =   1425
      Left            =   3840
      Picture         =   "frmmusic.frx":70EEE
      Top             =   4200
      Width           =   1905
   End
   Begin VB.Image Image3 
      Height          =   1845
      Left            =   0
      Picture         =   "frmmusic.frx":79DB0
      Top             =   3720
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   1425
      Left            =   1920
      Picture         =   "frmmusic.frx":85672
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   3840
      Picture         =   "frmmusic.frx":8E534
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This button allows the user to switch from the Search form
'to the subsequent Input form.
Private Sub cmdback_Click()
    frmsearch.Visible = False
    frmInput.Visible = True
End Sub

'This button allows the user to exit the program.
Private Sub cmdexit_Click()
    End
End Sub
'This button allows the user to search for information
'pertaining to a group that is input through the use of an inputbox.
Private Sub cmdgroup_Click()
    Dim Found As Boolean
    Dim Arraysize, Pos As Integer
    Found = False
    Dim X As String
    Arraysize = 20000
    Pos = 0
    X = InputBox("Enter Group's Name", "Group Name")
    Do Until Pos = SGCounter
        Pos = Pos + 1
        If LCase(X) = LCase(GroupArray(Pos)) Then
            picOutput.Print "Group",
            picOutput.Print GroupArray(Pos)
            picOutput.Print "CD", , "Genre"
            picOutput.Print CDArray(Pos), GenreArray(Pos)
            Found = True
        End If
    Loop
        If Found = False Then
            MsgBox "No Such Group", , "Nope!"
        End If
End Sub
'This button allows the user to search for songs that fit a certian
'mood that is input through the use of an inputbox.
Private Sub cmdmood_Click()
    Dim Found As Boolean
    Dim Arraysize, Pos As Integer
    Dim X As String
    Found = False
    Arraysize = 20000
    Pos = 0
    X = InputBox("Enter Your Mood", "Mood")
    Do Until Pos = SGCounter
        Pos = Pos + 1
        'MsgBox SongsArray(Pos)
        If LCase(X) = LCase(MoodArray(Pos)) Then
            picOutput.Print "Mood", , "Songs"
            picOutput.Print MoodArray(Pos), , SongsArray(Pos)
            picOutput.Print "CD", , "Genre"
            picOutput.Print CDArray(AlbumArray(Pos)), , GenreArray(AlbumArray(Pos))
            picOutput.Print "Group"
            picOutput.Print GroupArray(AlbumArray(Pos))
            Found = True
        End If
    Loop
    If Found = False Then
        MsgBox "No Such Mood", , "Nope!"
    End If
    

End Sub
'This button allows the user to search for information pertaining
'to any single song withing the text file CDs.txt by inputing
'the name of the desired song into an inputbox.
Private Sub cmdsong_Click()
    Dim Found As Boolean
    Dim Arraysize, Pos, A As Integer
    Dim X, I As String
    Arraysize = 20000
    Pos = 0
    Found = False
    X = InputBox("Enter Your Song", "Song")
    Do Until Pos = Arraysize Or Found = True
        Pos = Pos + 1
        If LCase(X) = LCase(SongsArray(Pos)) Then
            picOutput.Print "Song", , "Mood"
            picOutput.Print SongsArray(Pos), MoodArray(Pos)
            picOutput.Print "CD", , "Genre"
            picOutput.Print CDArray(AlbumArray(Pos)), GenreArray(AlbumArray(Pos))
            picOutput.Print "Group"
            picOutput.Print GroupArray(AlbumArray(Pos))
            Found = True
        End If
    Loop
        If Found = False Then
            MsgBox "No Such Song", , "Nope!"
        End If
End Sub
