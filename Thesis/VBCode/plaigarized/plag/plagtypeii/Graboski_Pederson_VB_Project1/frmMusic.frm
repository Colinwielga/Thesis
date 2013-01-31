VERSION 5.00
Begin VB.Form ggg
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain
      Caption         =   "Go to Main Menu"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      TabIndex        =   21
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext
      Caption         =   "Go to Next Slide"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11640
      TabIndex        =   20
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "Myriad Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7200
      TabIndex        =   19
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label Label11
      Alignment       =   2  'Center
      Caption         =   "Jared Allen"
      BeginProperty Font
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      TabIndex        =   18
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label10
      Alignment       =   2  'Center
      Caption         =   "Jermichael Finley"
      BeginProperty Font
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      TabIndex        =   17
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label9
      Alignment       =   2  'Center
      Caption         =   "Greg Jennings"
      BeginProperty Font
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label8
      Alignment       =   2  'Center
      Caption         =   "Donald Driver"
      BeginProperty Font
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label7
      Alignment       =   2  'Center
      Caption         =   "James Jones"
      BeginProperty Font
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   14
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label6
      Alignment       =   2  'Center
      Caption         =   "Spencer Havner"
      BeginProperty Font
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   13
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label5
      Alignment       =   2  'Center
      Caption         =   "Donald Lee"
      BeginProperty Font
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4
      Alignment       =   2  'Center
      Caption         =   "Jordy Nelson"
      BeginProperty Font
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   11
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label3
      Alignment       =   2  'Center
      Caption         =   "Double Click on the Player's Favorite Song Underneath their picture. P.S. We added a Viking in there too."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      TabIndex        =   10
      Top             =   2520
      Width           =   4095
   End
   Begin VB.OLE OLE8
      Class           =   "Package"
      Height          =   1215
      Left            =   12000
      OleObjectBlob   =   "ggg.frx":0000
      SourceDoc       =   "M:\CS130\Project\woman.mp3"
      TabIndex        =   9
      Top             =   6120
      Width           =   1335
   End
   Begin VB.OLE OLE7
      Class           =   "Package"
      Height          =   1335
      Left            =   7440
      OleObjectBlob   =   "ggg.frx":24018
      SourceDoc       =   "M:\CS130\Project\Solo.mp3"
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.OLE OLE6
      Class           =   "Package"
      Height          =   1335
      Left            =   5280
      OleObjectBlob   =   "ggg.frx":5D430
      SourceDoc       =   "M:\CS130\Project\popbottles.mp3"
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.OLE OLE5
      Class           =   "Package"
      Height          =   1215
      Left            =   9720
      OleObjectBlob   =   "ggg.frx":77648
      SourceDoc       =   "M:\CS130\Project\madeit.mp3"
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
   End
   Begin VB.OLE OLE4
      Class           =   "Package"
      Height          =   1215
      Left            =   4080
      OleObjectBlob   =   "ggg.frx":A5460
      SourceDoc       =   "M:\CS130\Project\gotmoney.mp3"
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
   End
   Begin VB.OLE OLE3
      Class           =   "Package"
      Height          =   1095
      Left            =   840
      OleObjectBlob   =   "ggg.frx":BCC78
      SourceDoc       =   "M:\CS130\Project\down.mp3"
      TabIndex        =   4
      Top             =   6600
      Width           =   1215
   End
   Begin VB.OLE OLE2
      Class           =   "Package"
      Height          =   1215
      Left            =   480
      OleObjectBlob   =   "ggg.frx":FA090
      SourceDoc       =   "M:\CS130\Project\Bedrock.mp3"
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.OLE OLE1
      Class           =   "Package"
      Height          =   1215
      Left            =   3000
      OleObjectBlob   =   "ggg.frx":133AA8
      SourceDoc       =   "M:\CS130\Project\baby.mp3"
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label2
      Alignment       =   2  'Center
      Caption         =   "Listen to The Receiver's Favorite Song"
      BeginProperty Font
         Name            =   "Lucida Handwriting"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7200
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label1
      Alignment       =   2  'Center
      Caption         =   "FUN FACTS TIME!"
      BeginProperty Font
         Name            =   "MS Mincho"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "ggg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get to know the Packers' Receivers
'ggg
'Brent Graboski
'2/15/10
'This form tells you the receivers favorite song

Private Sub poij_Click() 'this button shows the main menu
    aaa.Hide
    bbb.Show
    ccc.Hide
    ddd.Hide
    eee.Hide
    fff.Hide
    ggg.Hide
    hhh.Hide
End Sub

Private Sub qwer_Click() 'this button shows the next form
    aaa.Hide
    bbb.Hide
    ccc.Hide
    ddd.Hide
    eee.Hide
    fff.Hide
    ggg.Hide
    hhh.Show
End Sub

Private Sub zxcvb_Click() 'this button ends the program
    End
End Sub

