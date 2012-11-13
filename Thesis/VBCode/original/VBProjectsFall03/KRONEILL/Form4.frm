VERSION 5.00
Begin VB.Form AircraftPics 
   BackColor       =   &H00FF8080&
   Caption         =   "Boeing Aircraft Images"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Perpetua Titling MT"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   8175
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbxresults3 
      Height          =   6375
      Left            =   1560
      ScaleHeight     =   6315
      ScaleWidth      =   9075
      TabIndex        =   10
      Top             =   120
      Width           =   9135
   End
   Begin VB.CommandButton cmdmainmenu 
      Caption         =   "Return to Main Menu"
      Height          =   1095
      Left            =   8280
      TabIndex        =   8
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmd777300 
      Caption         =   "777-300"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmd767400 
      Caption         =   "767-400"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmd757300 
      Caption         =   "757-300"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmd757200 
      Caption         =   "757-200"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmd747400 
      Caption         =   "747-400"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmd737900 
      Caption         =   "737-900"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmd737800 
      Caption         =   "737-800"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmd717200 
      Caption         =   "717-200"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4920
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4920
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblauthor 
      BackColor       =   &H00FF8080&
      Caption         =   "VB Design by Kerry R. O'Neill 10/24/2003"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   7200
      Width           =   1695
   End
End
Attribute VB_Name = "AircraftPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The pupose of this form is to load and display a picture of
'each modern boeing aircraft a buyer may be interest in purchasing.
'This will give the user more familarization with Boeing's product
'line.

Option Explicit
Public strpath As String

Private Sub cmd717200_Click() 'loads a bitmap file from a folder for use in a picture box
    pbxresults3.Cls
    pbxresults3.Picture = LoadPicture(strpath & "717200.bmp")
End Sub

Private Sub cmd737800_Click() 'loads a bitmap file from a folder for use in a picture box
    pbxresults3.Cls
    pbxresults3.Picture = LoadPicture(strpath & "737800.bmp")
End Sub

Private Sub cmd737900_Click() 'loads a bitmap file from a folder for use in a picture box
    pbxresults3.Cls
    pbxresults3.Picture = LoadPicture(strpath & "737900.bmp")
End Sub

Private Sub cmd747400_Click() 'loads a bitmap file from a folder for use in a picture box
    pbxresults3.Cls
    pbxresults3.Picture = LoadPicture(strpath & "747400.bmp")
End Sub

Private Sub cmd757200_Click() 'loads a bitmap file from a folder for use in a picture box
    pbxresults3.Cls
    pbxresults3.Picture = LoadPicture(strpath & "757200.bmp")
End Sub

Private Sub cmd757300_Click() 'loads a bitmap file from a folder for use in a picture box
    pbxresults3.Cls
    pbxresults3.Picture = LoadPicture(strpath & "757300.bmp")
End Sub

Private Sub cmd767400_Click() 'loads a bitmap file from a folder for use in a picture box
    pbxresults3.Cls
    pbxresults3.Picture = LoadPicture(strpath & "767400.bmp")
End Sub

Private Sub cmd777300_Click() 'loads a bitmap file from a folder for use in a picture box
    pbxresults3.Cls
    pbxresults3.Picture = LoadPicture(strpath & "777300.bmp")
End Sub

Private Sub cmdmainmenu_Click() 'returns user to main menu
    AircraftPics.Hide
    MainMenu.Show
End Sub

Private Sub Form_Load() 'creates a strpath so the file can be opened after being moved to different folders
    strpath = "N:\CS130\handin\KRONEILL\"
End Sub
