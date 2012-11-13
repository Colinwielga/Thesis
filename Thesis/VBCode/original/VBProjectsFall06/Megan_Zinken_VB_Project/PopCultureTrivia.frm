VERSION 5.00
Begin VB.Form frmPopCultureTrivia 
   BackColor       =   &H80000007&
   Caption         =   "Pop Culture Trivia"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdProfile 
      Caption         =   "Profile"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13320
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   11880
      Picture         =   "Pop Culture Trivia.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1755
      TabIndex        =   12
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox picKillers 
      Height          =   1935
      Left            =   8160
      Picture         =   "Pop Culture Trivia.frx":AB4E
      ScaleHeight     =   1875
      ScaleWidth      =   2235
      TabIndex        =   11
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox picBritney 
      Height          =   1935
      Left            =   960
      Picture         =   "Pop Culture Trivia.frx":17F70
      ScaleHeight     =   1875
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtElton 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   11760
      TabIndex        =   9
      Text            =   "Elton John"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtKillers 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   8160
      TabIndex        =   8
      Text            =   "The Killers"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtTim 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Text            =   " Tim McGraw"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtBritney 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "Britney Spears"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.PictureBox picDisplay 
      Height          =   4455
      Left            =   1200
      ScaleHeight     =   4395
      ScaleWidth      =   9315
      TabIndex        =   5
      Top             =   5640
      Width           =   9375
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13320
      TabIndex        =   4
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCd 
      Caption         =   "Yearly Release"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13320
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdHits 
      Caption         =   "Album Info"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13320
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdGame 
      Caption         =   "Play Pop Culture Trivia"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2535
   End
   Begin VB.PictureBox picTim 
      Height          =   1935
      Left            =   4800
      Picture         =   "Pop Culture Trivia.frx":35472
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmPopCultureTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop Culture Trivia. (PopCultureTrivia.vbp)
'Form Name: frmPopCultureTrivia
'Author: Megan Zinken
'Date Written: November 2nd, 2006
'Form Objective: This form is used to play an MP3 for each artist and share information
                'about each artist.  With the profile button one can get information
                'about the artists such as hometown, age, and style of music.
                'With the Album Information button one can display information
                'about a certain artists album and year it was released.
                'With the album release button one can input a year and find
                'which artist, if any, has released an album that
                'year.
                'One can also press the Play! button to go to frmGame and play
                '"Pop Culture Trivia"
                
Private Sub cmdCd_Click()
 
    'This button prompts the user to input a year to display any album released in that year
    'Names, Albums, and Years are filed into an array
    'The user will input a year and it will display any album that corresponds to that year
    'If found = false there is not an album released in that year
    
    Dim InputYear As Integer
    Dim Singer(1 To 50) As String, Album(1 To 50) As String, Year(1 To 50) As Integer
    Dim Singers As String, Albums As String, Years As Integer
    Dim counter As Integer, size As Integer, found As Boolean
    picDisplay.Cls
    InputYear = InputBox("Which Year Would You Like Information For?")
    counter = 0
    Open App.Path & ("\SingerAlbumYear.txt") For Input As #1
        Do Until EOF(1)
            Input #1, Singers, Albums, Years
            counter = counter + 1
            Singer(counter) = Singers
            Album(counter) = Albums
            Year(counter) = Years
    Loop
    Close #1
    found = False
    picDisplay.Print "Name", Tab(30); ; "Album"; , Tab(75); ; "Release Year"
    picDisplay.Print "*******************************************************************************************************************************************"
    For size = 1 To counter
    If InputYear = Year(size) Then
        picDisplay.Print Singer(size), Tab(30); Album(size), Tab(75); Year(size)
        found = True
    End If
    Next size
    If found = False Then
        picDisplay.Cls
        picDisplay.Print ("There is Not an Album Released in That Year")
    End If
    
    
    
    
  
    
End Sub

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdGame_Click()
   'This button allows the user to play the game on frmGame
   'frmPopCultureTrivia disappears
   
   frmPopCultureTrivia.Visible = False
   frmGame.Visible = True

   End Sub
   
   
   

Private Sub cmdHits_Click()
  
    'This button is used to file Album information into an array
    'When clicked an input box is displayed and promps the user to input a name
    'It then displays information about the requested artist
    'If found = false there is not information about that artist
    
    Dim Singer(1 To 50) As String, Album(1 To 50) As String, Year(1 To 50) As Integer
    Dim Singers As String, Albums As String, Years As Integer
    Dim counter As Integer, InputName As String, pos As Integer, found As Boolean
    picDisplay.Cls
    InputName = InputBox("Which Artist Would You Like To Know More About?")
    counter = 0
    Open App.Path & ("\SingerAlbumYear.txt") For Input As #1
    Do Until EOF(1)
        Input #1, Singers, Albums, Years
        counter = counter + 1
        Singer(counter) = Singers
        Album(counter) = Albums
        Year(counter) = Years
    Loop
    Close #1
    found = False
    picDisplay.Print "Name", Tab(30); ; "Album"; , Tab(75); ; "Release Year"
    picDisplay.Print "*************************************************************************************************************************************************"
    For pos = 1 To counter
    If InputName = Singer(pos) Then
        found = True
        picDisplay.Print Singer(pos), Tab(30); Album(pos), Tab(80); Year(pos)
    End If
    Next pos
    If found = False Then
    MsgBox ("I'm Sorry We Do Not Have Information About That Artist")
    picDisplay.Cls
    End If
    
    
        
    
    
    
    
End Sub






Private Sub cmdProfile_Click()
    'This button is used to display profile information about a requested artist
    'Information is filed into an array
    'Via an input box the user will request an artist and corresponding information will be displayed
    'If found = false we do not have information on that artist
    
    Dim Artist(1 To 20) As String, Home(1 To 20) As String, Age(1 To 20) As String, Style(1 To 20) As String
    Dim sizes As Integer, Artists As String, Homes As String, Ages As String, Styles As String
    Dim Profiles As String, found As Boolean, pos As Integer
    Profiles = InputBox("Which Artist would you like to learn more about?")
    picDisplay.Cls
    Open App.Path & ("\Profile.txt") For Input As #2
    Do Until EOF(2)
        Input #2, Artists, Homes, Ages, Styles
        sizes = sizes + 1
        Artist(sizes) = Artists
        Home(sizes) = Homes
        Age(sizes) = Ages
        Style(sizes) = Styles
    Loop
    Close #2
    found = False
    For pos = 1 To sizes
    If Profiles = Artist(pos) Then
        found = True
        picDisplay.Print "Name: "; Artist(pos)
        picDisplay.Print "Home Town: "; Home(pos)
        picDisplay.Print "Age: "; Age(pos)
        picDisplay.Print "Stlye: "; Style(pos)
    End If
    Next pos
    If found = False Then
    picDisplay.Print "We do not have information on that artist"
    End If
    
        
        
        
    
    
End Sub
