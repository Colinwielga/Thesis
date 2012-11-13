VERSION 5.00
Begin VB.Form frmMusic 
   BackColor       =   &H8000000D&
   Caption         =   "Music"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form3"
   Picture         =   "frmMusic.frx":0000
   ScaleHeight     =   9885
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   8880
      Width           =   3015
   End
   Begin VB.CommandButton cmdSortS 
      Caption         =   "Sort By Song"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortA 
      Caption         =   "Sort By Artist"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   4800
      Width           =   5895
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show All Songs"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdGBQPlaylist 
      Caption         =   "The Good, the Bad and the Queen Playlist"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton cmdShinsPlaylist 
      Caption         =   "The Shins Playlist"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   3015
   End
   Begin VB.PictureBox picExtras 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   6360
      Picture         =   "frmMusic.frx":73D4
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   3060
   End
   Begin VB.PictureBox picGBQ 
      AutoSize        =   -1  'True
      Height          =   3090
      Left            =   3240
      Picture         =   "frmMusic.frx":97FE
      ScaleHeight     =   3030
      ScaleWidth      =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   3060
   End
   Begin VB.PictureBox picShins 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   120
      Picture         =   "frmMusic.frx":CF60
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   3060
   End
End
Attribute VB_Name = "frmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Artist(1 To 50) As String, Song(1 To 50) As String
'main music form






Private Sub cmdexit_Click()
frmMusic.Hide
End Sub

Private Sub cmdExtraPlaylist_Click()
    frmextras.Show  'shows the extras playlist
End Sub
'plays the shins playlist
Private Sub cmdShinsPlaylist_Click()
    Shell ("explorer.exe \\ad\homedir$\Students\A\a1smiles\My Documents\My Music\My Playlists\The Shins.wpl")
End Sub
'plays the GBQ playlist
Private Sub cmdGBQPlaylist_Click()
    Shell ("explorer.exe \\ad\homedir$\Students\A\a1smiles\My Documents\My Music\My Playlists\GBQ.wpl\")
End Sub

Private Sub cmdSortA_Click()
  Dim Pass As Integer
    Dim Pos As Integer
    Dim TempArtist As String
    Dim TempSong As String
    Dim TempAlbum As String
    
picResults.Cls                                          'sorts the MusicList file according to Artist
   Open App.Path & "\MusicList.txt" For Input As #3
    Do Until EOF(3)
    ctr = ctr + 1
    Input #3, Artist(ctr), Song(ctr)
    Loop
    Close #3
    For Pass = 1 To (ctr - 1)
        For Pos = 1 To (ctr - Pass)
            If Artist(Pos) > Artist(Pos + 1) Then
                TempArtist = Artist(Pos)
                Artist(Pos) = Artist(Pos + 1)
                Artist(Pos + 1) = TempArtist
                
                TempSong = Song(Pos)
                Song(Pos) = Song(Pos + 1)
                Song(Pos + 1) = TempSong
                
            End If
        Next Pos
    Next Pass
    picResults.Print "Artist"; Tab(35); "Song"
   picResults.Print "***********************************************************"
For Pos = 1 To ctr
       picResults.Print Artist(Pos); Tab(35), Song(Pos)
    Next Pos
End Sub

Private Sub Picture1_Click()
Dim Address(1 To 50) As String, ctr As Integer
    frmShinsAlbum.Show  'shows the shins album playlist
End Sub



Private Sub cmdShow_Click()

Dim ctr As Integer
picResults.Cls
Open App.Path & "\MusicList.txt" For Input As #3
    Do Until EOF(3)
        ctr = ctr + 1
        Input #3, Artist(ctr), Song(ctr)
    Loop
Close #3
   
   picResults.Print "Artist"; ""; Tab(35); "Song"
   picResults.Print "***************************************************************"
    
    For Pos = 1 To ctr
        picResults.Print Artist(Pos); Tab(35), Song(Pos)
    Next Pos
End Sub


Private Sub cmdSortS_Click()        'Sorts the MusicList according to song
  Dim Pass As Integer
    Dim Pos As Integer
    Dim TempArtist As String
    Dim TempSong As String
    Dim TempAlbum As String
    
picResults.Cls
   Open App.Path & "\MusicList.txt" For Input As #3
    Do Until EOF(3)
    ctr = ctr + 1
    Input #3, Artist(ctr), Song(ctr)
    Loop
    Close #3
    For Pass = 1 To (ctr - 1)
        For Pos = 1 To (ctr - Pass)
            If Song(Pos) > Song(Pos + 1) Then
                TempSong = Song(Pos)
                Song(Pos) = Song(Pos + 1)
                Song(Pos + 1) = TempSong
                
              TempArtist = Artist(Pos)
                Artist(Pos) = Artist(Pos + 1)
                Artist(Pos + 1) = TempArtist
                
            End If
        Next Pos
    Next Pass
    picResults.Print "Artist"; Tab(35); "Song"
   picResults.Print "***********************************************************"
For Pos = 1 To ctr
       picResults.Print Artist(Pos); Tab(35), Song(Pos)
    Next Pos
End Sub


Private Sub picExtras_Click()
    frmextras.Show
End Sub

Private Sub picGBQ_Click()
    frmGBQ.Show
End Sub

Private Sub picShins_Click()
    frmShinsAlbum.Show
End Sub
