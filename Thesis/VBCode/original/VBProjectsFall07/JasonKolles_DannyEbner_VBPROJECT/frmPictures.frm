VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H000080FF&
   Caption         =   "Player Pictures"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9000
      TabIndex        =   16
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton cmdBackHome 
      Caption         =   "Click Here to Go Back to the Home Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   15
      Top             =   7680
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000080FF&
      Height          =   6255
      Left            =   6720
      ScaleHeight     =   6195
      ScaleWidth      =   4515
      TabIndex        =   14
      Top             =   840
      Width           =   4575
   End
   Begin VB.CommandButton cmdLoadPics 
      Caption         =   "Load Pictures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowPicture 
      Caption         =   "Click Here to Show Picture"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      TabIndex        =   12
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtPlayerInput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblPicLabel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Player Pictures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblTypePlayer 
      BackColor       =   &H0080FFFF&
      Caption         =   "Enter the number that corresponds with the player on the list (Number Must Be 1-8)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2640
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblPicPlayers 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pictured Players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblNash 
      Caption         =   "8. Steve Nash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblMello 
      Caption         =   "7. Carmelo Anthony"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblChrisPaul 
      Caption         =   "6. Chris Paul"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblDuncan 
      Caption         =   "5. Tim Duncan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblAI 
      Caption         =   "4. Allen Iverson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblAmare 
      Caption         =   "3. Amare Stoudamire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblKobe 
      Caption         =   "2. Kobe Bryant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblPlayerPics 
      Caption         =   "1. Dirk Nowitzki"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim players(1 To 8) As String
Dim ctr As Integer

Private Sub cmdBackHome_Click()
    'brings the user back to the home page
frmHome.Show
frmPictures.Hide

End Sub

Private Sub cmdLoadPics_Click()

    'forces the user to push the load button first
    'disables the show picture option
cmdShowPicture.Enabled = True
cmdLoadPics.Enabled = False


    'open list of pictures from notepad file
    'dim the array as players above
    
Open App.Path & "\Players6.txt" For Input As #1


ctr = 0

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, players(ctr)
Loop
Close #1
End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub cmdShowPicture_Click()

    
 
Dim PlayerNum As Integer

    'ask user to put in a number that corresponds with the players from the list
    'make that number equal to PlayerNum
PlayerNum = txtPlayerInput.Text

    'have that number from the user determine which picture to show
    'loads the picture from the file and prints said picture
    

picResults.Picture = LoadPicture(App.Path & "\" & players(PlayerNum))
End Sub



