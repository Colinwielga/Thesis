VERSION 5.00
Begin VB.Form Gallery 
   BackColor       =   &H00000000&
   Caption         =   "Gallery"
   ClientHeight    =   12195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15870
   LinkTopic       =   "Form2"
   ScaleHeight     =   12195
   ScaleWidth      =   15870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Information"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   3
      Top             =   10320
      Width           =   2535
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next>>>"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      TabIndex        =   2
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<<Previous"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   1
      Top             =   8880
      Width           =   1815
   End
   Begin VB.PictureBox picPictures 
      Height          =   4095
      Left            =   5520
      ScaleHeight     =   4035
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label lblReal 
      BackColor       =   &H000000FF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   9
      Left            =   13200
      TabIndex        =   13
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label lblReal 
      BackColor       =   &H000000FF&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   8
      Left            =   13200
      TabIndex        =   12
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label lblReal 
      BackColor       =   &H000000FF&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   7
      Left            =   13200
      TabIndex        =   11
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblReal 
      BackColor       =   &H000000FF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   6
      Left            =   13200
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblReal 
      BackColor       =   &H000000FF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   5
      Left            =   13200
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblReal 
      BackColor       =   &H000000FF&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   4
      Left            =   13200
      TabIndex        =   8
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblReal 
      BackColor       =   &H00FF0000&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblReal 
      BackColor       =   &H00FF0000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblReal 
      BackColor       =   &H00FF0000&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   3
      Left            =   1080
      TabIndex        =   5
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label lblReal 
      BackColor       =   &H00FF0000&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   89.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "Gallery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form allows the user to toggle though some images from the players of real madrid on the pitch

'declare variables globally
Dim CTR As Integer
Dim path As String



Private Sub cmdBack_Click()
'Show and/or hide certian forms when "back" button is clicked
Information.Show
Gallery.Hide
OpenPage.Hide
PlayersStat.Hide
Trivia.Hide
Statistics.Hide
Form1.Hide
End Sub

Private Sub cmdNext_Click()
'begin if statement
If (CTR > 10) Then
CTR = 0 'set Counter to 0
End If
CTR = CTR + 1
picPictures.Picture = LoadPicture(path & CTR & ".jpg") 'load the picture file

End Sub

Private Sub cmdPrevious_Click()
If (CTR > 10) Then
CTR = 0 'set counter to 0
End If
CTR = CTR + 1
picPictures.Picture = LoadPicture(path & CTR & ".jpg")
End Sub

Private Sub Form_Load()
path = "M:\CS130\Images\gallery"
CTR = 0
End Sub
