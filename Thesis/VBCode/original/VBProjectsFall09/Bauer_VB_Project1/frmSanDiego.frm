VERSION 5.00
Begin VB.Form frmSanDiego 
   BackColor       =   &H80000013&
   Caption         =   "Form6"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form6"
   ScaleHeight     =   8220
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdWilds 
      BackColor       =   &H00FFFF00&
      Caption         =   "The Wilds Hotel"
      Height          =   1095
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmdHotelC 
      BackColor       =   &H00FFFF00&
      Caption         =   "The Castel Hotel"
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   3555
      Left            =   5400
      Picture         =   "frmSanDiego.frx":0000
      Top             =   360
      Width           =   4380
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   480
      Picture         =   "frmSanDiego.frx":32B3E
      Top             =   240
      Width           =   4200
   End
   Begin VB.Label cmdSanDiego 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "San Diego"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   6840
      Width           =   7935
   End
End
Attribute VB_Name = "frmSanDiego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'San Diego'
'creating an otption for people on where they would like to stay'
'Moves them between one frame to the next'
'October 14th 2009'
'Blake bauer'
'frm to hotel frm'
'Quit Button'
Private Sub cmdEnd_Click()
    End
End Sub

Private Sub cmdHotelC_Click()
    frmSanDiego.Hide
    frmHotel.Show
End Sub
'san diego frm to hotel frm'
Private Sub cmdWilds_Click()
    frmSanDiego.Hide
    frmHotel.Show
End Sub
