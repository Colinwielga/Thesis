VERSION 5.00
Begin VB.Form frmPlayerBio 
   BorderStyle     =   0  'None
   Caption         =   "PlayerBio"
   ClientHeight    =   15360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "PlayerBio.frx":0000
   ScaleHeight     =   15360
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPitchPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "See Pitchers Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdHitPage 
      BackColor       =   &H000000C0&
      Caption         =   "See Hitters Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000D&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H8000000D&
      Caption         =   "Find Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00800000&
      FillColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   3720
      ScaleHeight     =   2595
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Image picResults1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   960
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblLabel7 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Go to Pitchers Page:"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblLabel6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Go to Hitters Page:"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblLabel5 
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Search for a player to dsiplay:"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label lblLabel4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "End session:"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblLabel3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player Bio"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblLabel2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player Picture"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblLabel1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player Biographies of the Minnesota Twins (2009)"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   5175
   End
End
Attribute VB_Name = "frmPlayerBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
'Declare variables
Dim Found As Boolean, U As Integer, Ctr3 As Integer, Name As String, Player(1 To 50) As String, DOB(1 To 50) As Integer, POB(1 To 50) As String, BT(1 To 50) As String, HT(1 To 50) As Single, WT(1 To 50) As Integer, Debut(1 To 50) As Integer, CLG(1 To 50) As String
'initialize ctr
Ctr3 = 0
'open file
Open App.Path & "\TwinsRoster.txt" For Input As #1
'read file into an array
Do While Not EOF(1)
    Ctr3 = Ctr3 + 1
    Input #1, Player(Ctr3), DOB(Ctr3), POB(Ctr3), BT(Ctr3), HT(Ctr3), WT(Ctr3), Debut(Ctr3), CLG(Ctr3)
Loop
'close file
Close #1
'get user input
Name = InputBox("Enter the player's name on the 2009 Minnesota Twins that you would like to find (be sure to capitalize the first letter of the player's first and last name)")
U = 0
Found = False
'match stop search
Do While ((Not Found) And (U < Ctr3))
 U = U + 1
    If Name = Player(U) Then
    Found = True
    End If
Loop
picResults2.Cls
If (Not Found) Then
    picResults2.Print Name; " could not be found."
    Else
    picResults2.Print "1 player matched the search: "; Name; Chr(13); "Born: "; DOB(U); Chr(13); "Hometown: "; POB(U); Chr(13); "Bats/Throws: "; BT(U); Chr(13); "Height: "; HT(U); Chr(13); "Weight: "; WT(U); Chr(13); "Debut Year: "; Debut(U); Chr(13); "College: "; CLG(U)
End If
'incorporate a picture
Select Case Name
    Case Is = "Luis Ayala"
        picResults1.Picture = LoadPicture(App.Path & "\Luis_Ayala.jpg")
    Case Is = "Boof Bonser"
        picResults1.Picture = LoadPicture(App.Path & "\Boof_Bonser.jpg")
    Case Is = "Craig Breslow"
        picResults1.Picture = LoadPicture(App.Path & "\breslow.jpg")
    Case Is = "Jesse Crain"
        picResults1.Picture = LoadPicture(App.Path & "\Jesse_Crain.jpg")
    Case Is = "Brian Buscher"
        picResults1.Picture = LoadPicture(App.Path & "\Brian_Buscher.jpg")
    Case Is = "Kevin Slowey"
        picResults1.Picture = LoadPicture(App.Path & "\28KevinSlowey.jpg")
    Case Is = "Anthony Swarzak"
        picResults1.Picture = LoadPicture(App.Path & "\swarzak.jpg")
    Case Is = "Michael Cuddyer"
        picResults1.Picture = LoadPicture(App.Path & "\Cuddyer.jpg")
    Case Is = "Denard Span"
        picResults1.Picture = LoadPicture(App.Path & "\denard-span-1.jpg")
    Case Is = "Glen Perkins"
        picResults1.Picture = LoadPicture(App.Path & "\glen-perkins.jpg")
    Case Is = "Juan Morillo"
        picResults1.Picture = LoadPicture(App.Path & "\Juan Morillo.jpg")
    Case Is = "R.A. Dickey"
        picResults1.Picture = LoadPicture(App.Path & "\RA Dickey.jpg")
    Case Is = "Matt Guerrier"
        picResults1.Picture = LoadPicture(App.Path & "\Matt Guerrier.jpg")
    Case Is = "Jose Mijares"
        picResults1.Picture = LoadPicture(App.Path & "\Jose Mijares.jpg")
    Case Is = "Jose Morales"
        picResults1.Picture = LoadPicture(App.Path & "\morales-jose1.jpg")
    Case Is = "Mike Redmond"
        picResults1.Picture = LoadPicture(App.Path & "\Redmond.jpg")
    Case Is = "Delmon Young"
        picResults1.Picture = LoadPicture(App.Path & "\delmon1.jpg")
    Case Is = "Alexi Casilla"
        picResults1.Picture = LoadPicture(App.Path & "\Alexi.jpg")
    Case Is = "Carlos Gomez"
        picResults1.Picture = LoadPicture(App.Path & "\Carlos Gomez.jpg")
    Case Is = "Matt Tolbert"
        picResults1.Picture = LoadPicture(App.Path & "\tolbert.jpg")
    Case Is = "Joe Nathan"
        picResults1.Picture = LoadPicture(App.Path & "\JoeNathan.jpg")
    Case Is = "Pat Neshek"
        picResults1.Picture = LoadPicture(App.Path & "\Pat_Neshek.jpg")
    Case Is = "Fransisco Liriano"
        picResults1.Picture = LoadPicture(App.Path & "\Liriano.jpg")
    Case Is = "Nick Blackburn"
        picResults1.Picture = LoadPicture(App.Path & "\medium_blackburn.jpg")
    Case Is = "Scott Baker"
        picResults1.Picture = LoadPicture(App.Path & "\BakerTwins.jpg")
    Case Is = "Joe Crede"
        picResults1.Picture = LoadPicture(App.Path & "\Joe Crede.jpg")
    Case Is = "Brendan Harris"
        picResults1.Picture = LoadPicture(App.Path & "\Harris.jpg")
    Case Is = "Nick Punto"
        picResults1.Picture = LoadPicture(App.Path & "\Punto.jpg")
    Case Is = "Joe Mauer"
        picResults1.Picture = LoadPicture(App.Path & "\joe-mauer.jpg")
    Case Is = "Justin Morneau"
        picResults1.Picture = LoadPicture(App.Path & "\morneau.jpg")
    Case Is = "Jason Kubel"
        picResults1.Picture = LoadPicture(App.Path & "\Kubel.jpg")
    Case Is = "Phil Humber"
        picResults1.Picture = LoadPicture(App.Path & "\Phil Humber.jpg")
    Case Is = "Bobby Keppel"
        picResults1.Picture = LoadPicture(App.Path & "\Keppel.jpg")
    Case Is = "Ron Mahay"
        picResults1.Picture = LoadPicture(App.Path & "\Ron Mahayl.jpg")
    Case Is = "Jeff Manship"
        picResults1.Picture = LoadPicture(App.Path & "\Manship.jpg")
    Case Is = "Kevin Mulvey"
        picResults1.Picture = LoadPicture(App.Path & "\Mulvey.jpg")
    Case Is = "Carl Pavano"
        picResults1.Picture = LoadPicture(App.Path & "\CarlPavano.jpg")
    Case Is = "Jon Rauch"
        picResults1.Picture = LoadPicture(App.Path & "\Rauch.jpg")
    Case Is = "Orlando Cabrera"
        picResults1.Picture = LoadPicture(App.Path & "\orlando-cabrera.jpg")
    Case Is = "Justin Huber"
        picResults1.Picture = LoadPicture(App.Path & "\JustinHuber.jpg")
    Case Is = "Armando Gabino"
        picResults1.Picture = LoadPicture(App.Path & "\Gabino.jpg")
    Case Is = "Brian Duensing"
        picResults1.Picture = LoadPicture(App.Path & "\brian-duensing.jpg")
    Case Else
        picResults1.Picture = LoadPicture(App.Path & "\Ghost20Lady.jpg")
End Select
End Sub

Private Sub cmdHitPage_Click()
'show/hide form
frmHitPage.Show
frmPlayerBio.Hide
End Sub

Private Sub cmdPitchPage_Click()
'show/hidde form
frmPitchPage.Show
frmPlayerBio.Hide
End Sub

Private Sub cmdQuit_Click()
'quit
End
End Sub
