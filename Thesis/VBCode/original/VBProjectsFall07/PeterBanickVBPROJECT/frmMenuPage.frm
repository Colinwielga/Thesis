VERSION 5.00
Begin VB.Form frmMenuPage 
   Caption         =   "You Decide..."
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   DrawMode        =   1  'Blackness
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   60
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdCitations 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   13080
      Picture         =   "frmMenuPage.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmdPolls 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   13080
      Picture         =   "frmMenuPage.frx":7FA5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCareerStats 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   13080
      Picture         =   "frmMenuPage.frx":10204
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   13080
      Picture         =   "frmMenuPage.frx":18900
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton cmdBio 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   13080
      Picture         =   "frmMenuPage.frx":20F62
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "What's the Penalty?..."
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Poplar Std"
         Size            =   90
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   12615
   End
   Begin VB.Image picBackground 
      DragMode        =   1  'Automatic
      Height          =   21600
      Left            =   -6960
      Picture         =   "frmMenuPage.frx":29064
      Top             =   -120
      Width           =   38970
   End
End
Attribute VB_Name = "frmMenuPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBio_Click()
    'shows form bio for user to read about Rose
    frmBio.Show
End Sub

Private Sub cmdCareerStats_Click()
    'shows form CareerStats to see Rose's statistics
    frmCareerStats.Show
End Sub

Private Sub cmdCitations_Click()
    'shows form citations for user to see sources used in this project
    frmCitations.Show
End Sub

Private Sub cmdPolls_Click()
    'shows form Polls for user to take a survey and see the public opinion
    frmPolls.Show
End Sub

Private Sub cmdQuit_Click()
    'quits the entire program
    End
End Sub
