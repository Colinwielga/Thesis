VERSION 5.00
Begin VB.Form frmMiami 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form5"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form5"
   ScaleHeight     =   8160
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmdExclusive 
      BackColor       =   &H00FFC0FF&
      Caption         =   "This Place is so Exclusive It Does Not Have a Name!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton cmdColony 
      BackColor       =   &H00FFC0FF&
      Caption         =   "The Colony Hotel: Where the Rich and Beautiful Stay."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   2865
      Left            =   4080
      Picture         =   "frmMiami.frx":0000
      Top             =   4680
      Width           =   5865
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   360
      Picture         =   "frmMiami.frx":36DAA
      Top             =   120
      Width           =   5730
   End
   Begin VB.Label lblMiami 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Miami"
      BeginProperty Font 
         Name            =   "Myriad Web Pro Condensed"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      TabIndex        =   0
      Top             =   3000
      Width           =   6015
   End
End
Attribute VB_Name = "frmMiami"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'creating an otption for people on where they would like to stay'
'Moves them between one frame to the next'
'October 14th 2009'
'blake bauer'
'maimi to hotel'
Private Sub cmdColony_Click()
    frmMiami.Hide
    frmHotel.Show
End Sub
'miami to hotel'
Private Sub cmdExclusive_Click()
    frmMiami.Hide
    frmHotel.Show
End Sub
'Quit Button'
Private Sub cmdQuit1_Click()
    End
End Sub
