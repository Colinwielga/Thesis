VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00400000&
   Caption         =   "Twins Territory"
   ClientHeight    =   7905
   ClientLeft      =   4065
   ClientTop       =   1725
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   8085
   Begin VB.CommandButton cmdWorksCited 
      BackColor       =   &H000000C0&
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdShop 
      BackColor       =   &H000000C0&
      Caption         =   "Shopping!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdTrivia 
      BackColor       =   &H000000C0&
      Caption         =   "Twins Trivia"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H000000C0&
      Caption         =   "Name That Twin"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdFavoritePlayer 
      BackColor       =   &H000000C0&
      Caption         =   "Who's Your Favorite Twin?"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdRoster 
      BackColor       =   &H000000C0&
      Caption         =   "Roster"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdLineup 
      BackColor       =   &H000000C0&
      Caption         =   "Starting Lineup"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   7215
      Left            =   1800
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   7155
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H000000C0&
         Caption         =   "Exit Twins Territory"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End 'end program
End Sub

Private Sub cmdFavoritePlayer_Click()
    frmMain.Hide 'hides main form
    frmFavTwin.Show 'shows favorite Twin form
    'prompt instructions
    MsgBox "Click Load Names and then type the number corresponding to your favorite player on the list in the text box and click OK", , "Instructions"
End Sub

Private Sub cmdLineup_Click()
    frmMain.Hide 'hides main form
    frmStarting.Show 'shows Starting form
    'prompt instructions
    MsgBox "Click on a postion to see the best starter for that position.(In my opinion, at least)", , "Starting Lineup"
End Sub

Private Sub cmdName_Click()
    frmMain.Hide 'hides main form
    frmNameTwin.Show 'shows Name Twin form
End Sub

Private Sub cmdRoster_Click()
    frmMain.Hide 'hides main form
    frmRoster.Show 'shows Roster form
End Sub

Private Sub cmdShop_Click()
    frmMain.Hide 'hides main form
    frmShopping.Show 'shows shopping form
    'prompt instructions
    MsgBox "Click on Start Shopping to begin selecting items. Click on the item picture and enter a size. When you have finished shopping, click Done Shopping to complete your purchase.", , "Shopping" 'prompts instructions
  
End Sub

Private Sub cmdTrivia_Click()
    frmMain.Hide 'hides main form
    frmTrivia.Show 'shows trivia form
    'prompt instructions
    MsgBox "In the text box next to each quesiton, enter the captial letter that corresponds to the correct answer for that question. When you have answered every question, click Finish", , "Instructions" 'prompts instructions
End Sub

Private Sub cmdWorksCited_Click()
    frmMain.Hide 'hides main form
    frmWorksCited.Show 'shows works cited form
End Sub
