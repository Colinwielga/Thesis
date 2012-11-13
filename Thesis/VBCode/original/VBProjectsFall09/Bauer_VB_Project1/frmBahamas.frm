VERSION 5.00
Begin VB.Form frmBahamas 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form2"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Quit"
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdHilton 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Hilton of the Bahamas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton cmdAtlantis 
      BackColor       =   &H00FFFFC0&
      Caption         =   "The Atlantis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   3060
      Left            =   5040
      Picture         =   "frmBahamas.frx":0000
      Top             =   3000
      Width           =   4350
   End
   Begin VB.Image Image1 
      Height          =   3195
      Left            =   360
      Picture         =   "frmBahamas.frx":2B722
      Top             =   3000
      Width           =   4245
   End
   Begin VB.Label lblWhereToStay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Where Would You Like To Stay?"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   8415
   End
   Begin VB.Label lblBahamas 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "The Bahamas"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "frmBahamas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'bahamas Page'
'October 15th 2009'
'This is another reansition page to the hotel page
'it is a show and hide page with pictures and command buttons'
'Blake bauer'


'once again changing to the hotel page'
Private Sub cmdAtlantis_Click()
    frmBahamas.Hide
    frmHotel.Show
End Sub
'Quit Button'
Private Sub cmdEnd_Click()
    End
End Sub
'once again changing to the hotel page'
Private Sub cmdHilton_Click()
    frmBahamas.Hide
    frmHotel.Show
End Sub
