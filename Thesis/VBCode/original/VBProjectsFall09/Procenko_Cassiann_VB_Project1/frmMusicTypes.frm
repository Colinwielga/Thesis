VERSION 5.00
Begin VB.Form frmMusicTypes 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Music Genres"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdNature 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sounds of Nature"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdCountry 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdRock 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rock and Roll"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      MaskColor       =   &H00FF0000&
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdBlues 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Blues"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   1620
      Left            =   7080
      Picture         =   "frmMusicTypes.frx":0000
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   780
      Left            =   6840
      Picture         =   "frmMusicTypes.frx":0964
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1620
      Left            =   120
      Picture         =   "frmMusicTypes.frx":120E
      Top             =   2640
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   120
      Picture         =   "frmMusicTypes.frx":4895
      Top             =   1440
      Width           =   2145
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please click on your favorite music type to see musicians/music of that genre."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmMusicTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Music VB Project by Cassiann Procenko
'Form Name is frmMusicTypes
'Date written 10/16/2009
'Purpose of this form is to display the different music genres the viewer can choose from

Private Sub cmdBlues_Click()
'show and hide forms
    frmBlues.Show
    frmMusicTypes.Hide
End Sub
Private Sub cmdCountry_Click()
'show and hide forms
    frmCountry1.Show
    frmMusicTypes.Hide
End Sub

Private Sub cmdNature_Click()
'show and hide forms
    frmNature.Show
    frmMusicTypes.Hide
End Sub

Private Sub cmdRock_Click()
'show and hide forms
    frmRockandRoll.Show
    frmMusicTypes.Hide
End Sub

Private Sub cmdReturn_Click()
'show and hide forms
    frmMusicHall.Show
    frmMusicTypes.Hide
End Sub

Private Sub cmdQuit_Click()
'show and hide forms
    frmLeave.Show
    frmMusicTypes.Hide
End Sub

