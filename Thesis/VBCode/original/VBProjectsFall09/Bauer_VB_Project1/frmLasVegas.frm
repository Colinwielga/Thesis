VERSION 5.00
Begin VB.Form frmLasVegas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form7"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form7"
   ScaleHeight     =   8310
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdParis 
      BackColor       =   &H80000013&
      Caption         =   "Paris in Las Vegas"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdMGM 
      BackColor       =   &H80000013&
      Caption         =   "MGM Grand"
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
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   3165
      Left            =   5760
      Picture         =   "frmLasVegas.frx":0000
      Top             =   3600
      Width           =   4350
   End
   Begin VB.Image Image3 
      Height          =   3240
      Left            =   360
      Picture         =   "frmLasVegas.frx":2CEFA
      Top             =   3480
      Width           =   4305
   End
   Begin VB.Label lblStay 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Where Would You like To Stay?"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   1680
      Left            =   8400
      Picture         =   "frmLasVegas.frx":5A83C
      Top             =   1440
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   1590
      Left            =   0
      Picture         =   "frmLasVegas.frx":66E3E
      Top             =   1560
      Width           =   2205
   End
   Begin VB.Label lblLasVegas 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Las Vegas"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmLasVegas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Welcome to las vegas page'
'Option of two hotels'
'I'm going to go from one page to the pricing page'
'October 14th'
'Blake Bauer'
Option Explicit

'Quit Button'
Private Sub cmdEnd2_Click()
    End
End Sub

Private Sub cmdMGM_Click()
    frmLasVegas.Hide
    frmHotel.Show
End Sub

Private Sub cmdParis_Click()
    frmLasVegas.Hide
    frmHotel.Show
End Sub
