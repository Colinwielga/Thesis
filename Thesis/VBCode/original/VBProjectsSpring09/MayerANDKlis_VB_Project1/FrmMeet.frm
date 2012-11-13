VERSION 5.00
Begin VB.Form FrmMeet 
   BackColor       =   &H00008000&
   Caption         =   "Let's Meet the Minnesota Twins Starting Line up"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   FillColor       =   &H80000004&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Main Form"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   2055
   End
   Begin VB.PictureBox picName 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   4320
      Width           =   2295
   End
   Begin VB.PictureBox picPlayer 
      Height          =   2655
      Left            =   4920
      ScaleHeight     =   2595
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton CmdRight 
      Caption         =   "Right Field"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton CmdCenter 
      Caption         =   "Center Field"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton CmdLeft 
      Caption         =   "Left Field"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton CmdThird 
      Caption         =   "Third Base"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton CmdShort 
      Caption         =   "Shortstop"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton CmdCatcher 
      Caption         =   "Catcher"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   2
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "First Base"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      TabIndex        =   1
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton CmdSecond 
      Caption         =   "Second Base"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      TabIndex        =   0
      Top             =   2880
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMeet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Minnesota Twins
'FrmMeet
'Jake Klis and Sarah Mayer
'Written on 3/21/09
'This form is used to display the players name and picture when the button of their
'corresponding field position is clicked
Private Sub Cmdback_Click()
FrmMeet.Hide
FrmMain.Show
End Sub
'When clicked Joe Mauer's picture and name will be displayed
Private Sub CmdCatcher_Click()
picName.Cls
picPlayer.Cls
picName.Print "Joe Mauer"
picPlayer.Picture = LoadPicture(App.Path & "\Mauer.jpg")
End Sub
'When clicked Justin Morneau's picture and name will be displayed
Private Sub CmdFirst_Click()
picName.Cls
picPlayer.Cls
picName.Print "Justin Morneau"
picPlayer.Picture = LoadPicture(App.Path & "\Morneau.jpg")
End Sub
'When clicked Alexi Cassila's picture and name will be displayed
Private Sub CmdSecond_Click()
picName.Cls
picPlayer.Cls
picName.Print "Alexi Casilla"
picPlayer.Picture = LoadPicture(App.Path & "\Casilla.jpg")
End Sub
'When clicked Joe Crede's picture and name will be displayed
Private Sub CmdThird_Click()
picName.Cls
picPlayer.Cls
picName.Print "Joe Crede"
picPlayer.Picture = LoadPicture(App.Path & "\Crede.jpg")
End Sub
'When clicked Nick Punto's picture and name will be displayed
Private Sub CmdShort_Click()
picName.Cls
picPlayer.Cls
picName.Print "Nick Punto"
picPlayer.Picture = LoadPicture(App.Path & "\Punto.jpg")
End Sub
'When clicked Delmon Young's picture and name will be displayed
Private Sub CmdLeft_Click()
picName.Cls
picPlayer.Cls
picName.Print "Delmon Young"
picPlayer.Picture = LoadPicture(App.Path & "\Young.jpg")
End Sub
'When clicked Michael Cuddyer's picture and name will be displayed
Private Sub CmdRight_Click()
picName.Cls
picPlayer.Cls
picName.Print "Michael Cuddyer"
picPlayer.Picture = LoadPicture(App.Path & "\Cuddyer.jpg")
End Sub
Private Sub CmdCenter_Click()
'When clicked Denard Span's picture and name will be displayed
picName.Cls
picPlayer.Cls
picName.Print "Denard Span"
picPlayer.Picture = LoadPicture(App.Path & "\Span.jpg")
End Sub
