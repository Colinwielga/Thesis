VERSION 5.00
Begin VB.Form frmDublin 
   BackColor       =   &H0000C000&
   Caption         =   "Dublin"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   16950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCivic 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   480
      ScaleHeight     =   3735
      ScaleWidth      =   5295
      TabIndex        =   6
      Top             =   1200
      Width           =   5295
   End
   Begin VB.PictureBox picGael 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   3720
      ScaleHeight     =   1935
      ScaleWidth      =   2415
      TabIndex        =   5
      Top             =   5280
      Width           =   2415
   End
   Begin VB.PictureBox picLibrary 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2775
      ScaleWidth      =   2775
      TabIndex        =   4
      Top             =   5760
      Width           =   2775
   End
   Begin VB.PictureBox picPub 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3960
      ScaleHeight     =   1335
      ScaleWidth      =   1935
      TabIndex        =   3
      Top             =   7560
      Width           =   1935
   End
   Begin VB.PictureBox picParade 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   6840
      ScaleHeight     =   7095
      ScaleWidth      =   9495
      TabIndex        =   2
      Top             =   1080
      Width           =   9495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Form"
      Height          =   855
      Left            =   11760
      TabIndex        =   1
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   13800
      TabIndex        =   0
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Tim Johnson   3/20    Just to show a bit about my home town"
      Height          =   615
      Left            =   13080
      TabIndex        =   8
      Top             =   9360
      Width           =   2895
   End
   Begin VB.Label lblDublinTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Dublin, CA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      TabIndex        =   7
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmDublin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBack_Click()     'Goes back to Main Form

frmMain.Show                    'Goes back to Main Form
frmDublin.Hide

End Sub

Private Sub cmdQuit_Click()     'Ends program where you are
    End                         'Ends program where you are
End Sub

Private Sub Form_Load()         'Puts ups pictures to improve form appearance

picParade.Picture = LoadPicture(App.Path & "\" & dublinpix(1))
picCivic.Picture = LoadPicture(App.Path & "\" & dublinpix(2))
picLibrary.Picture = LoadPicture(App.Path & "\" & dublinpix(3))
picGael.Picture = LoadPicture(App.Path & "\" & dublinpix(4))
picPub.Picture = LoadPicture(App.Path & "\" & dublinpix(5))

End Sub
