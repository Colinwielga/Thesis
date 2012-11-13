VERSION 5.00
Begin VB.Form frmPics 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   11610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   11610
   ScaleWidth      =   14205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H0080FFFF&
      Caption         =   "Go to Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8880
      Width           =   2175
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0080FFFF&
      Caption         =   "Go To Next Slide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10080
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   3960
      Picture         =   "frmPics.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   8355
      TabIndex        =   9
      Top             =   3240
      Width           =   8415
   End
   Begin VB.CommandButton cmdHavner 
      BackColor       =   &H0080FFFF&
      Caption         =   "Spencer Havner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdJames 
      BackColor       =   &H0080FFFF&
      Caption         =   "James Jones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdGreg 
      BackColor       =   &H0080FFFF&
      Caption         =   "Gregory Jennings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdJordy 
      BackColor       =   &H0080FFFF&
      Caption         =   "Jordy Nelson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdDLee 
      BackColor       =   &H0080FFFF&
      Caption         =   "Donald Lee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdJF 
      BackColor       =   &H0080FFFF&
      Caption         =   "Jermichael Finley"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdDonald 
      BackColor       =   &H0080FFFF&
      Caption         =   "Donald Driver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   1335
      Left            =   3720
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape Shape13 
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   3480
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Meet The Best Receiving Core in the NFL!!"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   2775
      Left            =   5160
      TabIndex        =   11
      Top             =   240
      Width           =   7455
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H0000FFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   11055
      Left            =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Get to know the Packers' Receivers
'frmPics
'Sam Pederson
'2/23/10
'This form shows the pictures of all the receivers of the Packers

Private Sub cmdDLee_Click() 'this button loads the picture of Donald Lee
    picResults.Picture = LoadPicture(App.Path & "\DLee.jpg")
End Sub

Private Sub cmdDonald_Click() 'this button loads the picture of Donald Driver
    picResults.Picture = LoadPicture(App.Path & "\Donald.jpg")
End Sub

Private Sub cmdGreg_Click() 'this button loads the picture of Greg Jennings
    picResults.Picture = LoadPicture(App.Path & "\Greg.jpg")
End Sub

Private Sub cmdHavner_Click() 'this button loads the picture of Spencer Havner
    picResults.Picture = LoadPicture(App.Path & "\Havner.jpg")
End Sub

Private Sub cmdJames_Click() 'this button loads the picture of James Jones
    picResults.Picture = LoadPicture(App.Path & "\James.jpg")
End Sub

Private Sub cmdJF_Click() 'this button loads the picture of Jermichael Finley
    picResults.Picture = LoadPicture(App.Path & "\JF.jpg")
End Sub

Private Sub cmdJordy_Click() 'this button loads the picture of Jordy Nelson
    picResults.Picture = LoadPicture(App.Path & "\Jordy.jpg")
End Sub

Private Sub cmdMain_Click() 'this button takes you to the menu form
    frmWelcome.Hide
    frmMenu.Show
    frmPoll.Hide
    frmData.Hide
    frmSwap.Hide
    frmPics.Hide
    frmMusic.Hide
    frmLast.Hide
End Sub

Private Sub cmdNext_Click() 'this button takes you to the next form
    frmWelcome.Hide
    frmPoll.Hide
    frmData.Hide
    frmSwap.Hide
    frmPics.Hide
    frmMusic.Show
    frmLast.Hide
End Sub

Private Sub cmdQuit_Click() 'this button ends the program
    End
End Sub
