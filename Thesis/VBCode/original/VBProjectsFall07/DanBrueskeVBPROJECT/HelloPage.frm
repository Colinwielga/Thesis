VERSION 5.00
Begin VB.Form FrmWelcome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCite 
      Caption         =   "Citations"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8280
      TabIndex        =   4
      Top             =   9960
      Width           =   3975
   End
   Begin VB.CommandButton cmdStore 
      Caption         =   "Shopping Store"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8280
      TabIndex        =   3
      Top             =   5880
      Width           =   3975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000040C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8280
      MaskColor       =   &H000040C0&
      TabIndex        =   2
      Top             =   7920
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000040C0&
      Caption         =   "Click To Play"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8280
      TabIndex        =   1
      Top             =   3840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Creator:  Dan Brueske"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   5
      Top             =   11880
      Width           =   2775
   End
   Begin VB.Label lblBegin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Let's Play Some Sports Trivia!!"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   0
      Top             =   2400
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   17100
      Left            =   1440
      Picture         =   "HelloPage.frx":0000
      Top             =   -240
      Width           =   17280
   End
End
Attribute VB_Name = "FrmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCite_Click()
    'This transfers forms. It transfers the welcome form to the cite form.
FrmWelcome.Hide
FrmCite.Show

End Sub

    'This form is the first form the user sees. It asks for their name in a message box and stores it globally.

Private Sub cmdQuit_Click()
    'Ends the program.
End
End Sub

Private Sub CmdStore_Click()
    'It transfers the forms.  From the welcome form to the store form.
FrmWelcome.Hide
FrmStore.Show

End Sub

Private Sub Command1_Click()
    'It globally stores the user name through an input box to be used later.  It also transfers the welcome form to the sports form.
UserName = InputBox("What is Your Name?")
FrmSports.Show
FrmWelcome.Hide

End Sub
