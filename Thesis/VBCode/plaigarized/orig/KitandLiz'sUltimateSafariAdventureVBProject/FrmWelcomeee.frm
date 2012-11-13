VERSION 5.00
Begin VB.Form FrmWelcome 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Eras Bold ITC"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11040
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRateService 
      BackColor       =   &H00404080&
      Caption         =   "Rate our Service!"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10320
      Width           =   2535
   End
   Begin VB.CommandButton cmdSlideshowandmusic 
      BackColor       =   &H00800080&
      Caption         =   "See Picutes! Hear Music!"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9600
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00800080&
      Caption         =   "Return to Civilization"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   2655
   End
   Begin VB.CommandButton cmdDesert 
      BackColor       =   &H0000C0C0&
      Caption         =   "Brave the Desert"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdRiver 
      BackColor       =   &H00808000&
      Caption         =   "Go to the Rushing River "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton cmdJungle 
      BackColor       =   &H00008000&
      Caption         =   "Enter the Jungle Oasis"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      MaskColor       =   &H00400040&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      FillColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   1200
      Picture         =   "FrmWelcome.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   2760
      Width           =   5895
   End
   Begin VB.Label lblwelcome2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   $"FrmWelcome.frx":24CE
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   7560
      Width           =   6375
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Karibu tena, Welcome"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   7080
      Width           =   4575
   End
   Begin VB.Label lblsafari 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "The Great Safari Adventure"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "FrmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Great Safari Adventure
'Frm Welcome
'Kit and Liz Chambers
'February 18th 2010
'Objective: The Purpose of this form is too
    'be a welcome screen
     'Bring user to other forms
     'Show a pictutre
     'Quit Button
     
Private Sub cmdDesert_Click()
    'gets information from inputbox
    UserName = InputBox("Please Enter Your Name:", "Welcome!")
    FrmWelcome.Hide 'hides Welcome page from user
    FrmTheDesert.Show 'shows desert page to user
End Sub

Private Sub cmdJungle_Click()
    'gets information from input box
    UserName = InputBox("What's your name?", "Welcome!")
    FrmWelcome.Hide 'hides Welcome page from user
    frmTheJungle.Show 'shows main page to user
    MsgBox "Welcome to the Jungle " & UserName & " We've got fun and games!", , "Enter the Jungle."
End Sub

Private Sub cmdQuit_Click()
MsgBox "Enjoy your Day!", , "Exit"
End
'Quits the program

End Sub

Private Sub cmdRateService_Click()
FrmWelcome.Hide 'hides the welcome form
frmRateService.Show 'shows the form to rate the service
End Sub

Private Sub cmdRiver_Click()
'gets information from input box
UserName = InputBox("What's your name?", "Welcome!")
FrmWelcome.Hide 'hides Welcome page from user
FrmTheRiver.Show 'shows river page
MsgBox "Hi, " & UserName & "! Adventure is waiting just around the river bend!", , "The River Awaits."

End Sub

Private Sub cmdSlideshowandmusic_Click()
FrmWelcome.Hide 'Hides the welcome form
frmSafariMusic.Show 'Shows music and slideshow form


End Sub

Private Sub picResults_Click()
picResults.Show africansafari.jpeg

End Sub
