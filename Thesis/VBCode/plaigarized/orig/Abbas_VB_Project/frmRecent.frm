VERSION 5.00
Begin VB.Form frmRecent 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000C0C0&
      Caption         =   "Exit the Program"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H0000C0C0&
      Caption         =   "Go Back to the Main Menu"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   3960
      ScaleHeight     =   6255
      ScaleWidth      =   6735
      TabIndex        =   10
      Top             =   240
      Width           =   6735
   End
   Begin VB.CommandButton cmdWeinke 
      BackColor       =   &H0000C0C0&
      Caption         =   "2000 - Chris Weinke"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCrouch 
      BackColor       =   &H0000C0C0&
      Caption         =   "2001 - Eric Crouch"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdPalmer 
      BackColor       =   &H0000C0C0&
      Caption         =   "2002 - Carson Palmer"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdWhite 
      BackColor       =   &H0000C0C0&
      Caption         =   "2003 - Jason White"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdLeinart 
      BackColor       =   &H0000C0C0&
      Caption         =   "2004 - Matt Leinart"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdBush 
      BackColor       =   &H0000C0C0&
      Caption         =   "2005 - Reggie Bush"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdSmith 
      BackColor       =   &H0000C0C0&
      Caption         =   "2006 - Troy Smith"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdTebow 
      BackColor       =   &H0000C0C0&
      Caption         =   "2007 - Tim Tebow"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdBradford 
      BackColor       =   &H0000C0C0&
      Caption         =   "2008 - Sam Bradford"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdIngram 
      BackColor       =   &H0000C0C0&
      Caption         =   "2009 - Mark Ingram"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmRecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Heisman Trophy
'frmRecent
'Kevin Abbas
'2-21-10
'Objective - To show pictures of the 10 most recent Heisman winners

Dim Year As Integer

Private Sub cmdBradford_Click() 'load and display Sam Bradford's picture
    picResults.Picture = LoadPicture(App.Path & "\Bradford.jpg")
End Sub

Private Sub cmdBush_Click()
    picResults.Picture = LoadPicture(App.Path & "\Bush.jpg")
End Sub

Private Sub cmdCrouch_Click()
    picResults.Picture = LoadPicture(App.Path & "\Crouch.jpg")
End Sub

Private Sub cmdExit_Click() 'end the program and thank the user
    MsgBox ("Hope you enjoyed learning about the Heisman, have a nice day!")
    End
End Sub

Private Sub cmdIngram_Click()
    picResults.Picture = LoadPicture(App.Path & "\MarkIngram.jpg")
End Sub

Private Sub cmdLeinart_Click()
    picResults.Picture = LoadPicture(App.Path & "\Leinart.jpg")
End Sub

Private Sub cmdMain_Click() 'bring the user back to the main menu
    frmWelcome.Show
    frmHistory.Hide
    frmWinners.Hide
    frmWhereNow.Hide
    frmRecent.Hide
End Sub

Private Sub cmdPalmer_Click()
    picResults.Picture = LoadPicture(App.Path & "\Palmer.jpg")
End Sub


Private Sub cmdSmith_Click()
    picResults.Picture = LoadPicture(App.Path & "\Smith.jpg")
End Sub

Private Sub cmdTebow_Click()
    picResults.Picture = LoadPicture(App.Path & "\Tebow.jpg")
End Sub

Private Sub cmdWeinke_Click()
    picResults.Picture = LoadPicture(App.Path & "\Weinke.jpg")
End Sub

Private Sub cmdWhite_Click()
    picResults.Picture = LoadPicture(App.Path & "\White.jpg")
End Sub

Private Sub Form_Load()
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
