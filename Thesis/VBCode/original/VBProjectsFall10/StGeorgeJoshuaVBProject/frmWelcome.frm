VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Welcome!"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13575
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13575
   Begin VB.CommandButton cmdEnter 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Enter Program"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label lblCitation 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.earthinpictures.com/world/italy/rome/colosseum_and_moon.html"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   5535
   End
   Begin VB.Label lblTitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bringing the Latin Language to life "
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   9600
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblCreator 
      BackColor       =   &H00000080&
      Caption         =   "Designed by Josh StGeorge Oct 2010"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   8280
      Width           =   9135
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00000080&
      Caption         =   "Lingua Vivens!:"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   27.75
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   9600
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Image imgColleseumTitle 
      Height          =   7200
      Left            =   0
      Picture         =   "frmWelcome.frx":0000
      Top             =   720
      Width           =   9600
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEnter_Click()
    'moves from welcome screen to the login screen
    frmWelcome.Hide
    frmPreLogin.Show
End Sub

