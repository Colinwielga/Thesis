VERSION 5.00
Begin VB.Form frmStarting 
   BackColor       =   &H00008000&
   Caption         =   "Starting Lineup"
   ClientHeight    =   7785
   ClientLeft      =   2955
   ClientTop       =   1890
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdCloser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Closer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdReliever 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reliever"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   2415
      TabIndex        =   12
      Top             =   6720
      Width           =   2415
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   1335
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit Twins Territory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to Twins Territory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdPitcher 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pitcher"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdLeft 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Left Field"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdRight 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Right Field"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdFourth 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Third Base"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdShort 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shortstop"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Base"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCenter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Center Field"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSecond 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Second Base"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdCatcher 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Catcher"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   8
      X1              =   1200
      X2              =   3600
      Y1              =   3240
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   8
      X1              =   6600
      X2              =   4320
      Y1              =   3240
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   8
      X1              =   1320
      X2              =   3120
      Y1              =   4200
      Y2              =   6120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   8
      X1              =   4680
      X2              =   6600
      Y1              =   6120
      Y2              =   4200
   End
End
Attribute VB_Name = "frmStarting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
'goes back to Title form
    frmStarting.Hide
    frmMain.Show
End Sub

Private Sub cmdCatcher_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Joe Mauer picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\JoeMauer.jpg")
    picName.Print "Joe Mauer"
End Sub

Private Sub cmdCenter_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\ToriiHunter.jpg")
    picName.Print "Torii Hunter"
End Sub

Private Sub cmdCloser_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\JoeNathan.jpg")
    picName.Print "Joe Nathan"
End Sub

Private Sub cmdFirst_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\JustinMorneau.jpg")
    picName.Print "Justin Morneau"
End Sub

Private Sub cmdFourth_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\NickPunto.jpg")
    picName.Print "Nick Punto"
End Sub

Private Sub cmdLeft_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\JasonTyner.jpg")
    picName.Print "Jason Tyner"
End Sub

Private Sub cmdPitcher_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\JohanSantana.jpg")
    picName.Print "Johan Santana"
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReliever_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\PatNeshek.jpg")
    picName.Print "Pat Neshek"
End Sub

Private Sub cmdRight_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\MichaelCuddyer.jpg")
    picName.Print "Michael Cuddyer"
End Sub

Private Sub cmdSecond_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\AlexiCasilla.jpg")
    picName.Print "Alexi Casilla"
End Sub

Private Sub cmdShort_Click()
'clear picture boxes
    picPlayer.Cls
    picName.Cls
'Displays Johan Santana picture and name
    picPlayer.Picture = LoadPicture(App.Path & "\projectphotos\JasonBartlett.jpg")
    picName.Print "Jason Bartlett"
End Sub
