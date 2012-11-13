VERSION 5.00
Begin VB.Form frmMainPage 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main Page"
   ClientHeight    =   9405
   ClientLeft      =   5385
   ClientTop       =   3540
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   15525
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   4440
      Picture         =   "frmMainPage.frx":0000
      ScaleHeight     =   3615
      ScaleWidth      =   5295
      TabIndex        =   4
      Top             =   840
      Width           =   5295
   End
   Begin VB.CommandButton cmdCreators 
      BackColor       =   &H8000000D&
      Caption         =   "Meet The Creators "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdRules 
      BackColor       =   &H000080FF&
      Caption         =   "Rules"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton cmdStartGame 
      BackColor       =   &H8000000D&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   " Quit "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2295
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Family Feud
'frmMainPage
'Colin Hall and Andre Blaine
'March 15
'This form will hide the Main Page form,
'will show various other forms such as: the Family Name form, the Creators form, and the Rules form,
'and will quit.

Private Sub cmdStartGame_Click()

    'This button will hide the Main Page form and will show the Family Name form,
    frmMainPage.Hide
    frmFamilyName.Show
    
End Sub

Private Sub cmdCreators_Click()

    'This button will hide the Main Page form and will show the Creators form.
    frmMainPage.Hide
    frmCreators.Show
    
End Sub

Private Sub cmdRules_Click()

    'This button will hide the Main Page form and will show the Rules form.
    frmMainPage.Hide
    frmRules.Show
    
End Sub

Private Sub cmdQuit_Click()

    'This button will end the Visual Basic Program.
    End
    
End Sub
