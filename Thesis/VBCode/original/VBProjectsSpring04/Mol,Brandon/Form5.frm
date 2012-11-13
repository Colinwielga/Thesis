VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H0000C000&
   Caption         =   "Touring"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form5"
   ScaleHeight     =   8670
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Click to return to the main menu"
      Height          =   1215
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   2175
   End
   Begin VB.PictureBox Picture7 
      Height          =   2055
      Left            =   4200
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   6
      Top             =   2640
      Width           =   3255
   End
   Begin VB.PictureBox Picture6 
      Height          =   2055
      Left            =   8160
      Picture         =   "Form5.frx":5CAA
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   5280
      Width           =   3255
   End
   Begin VB.PictureBox Picture5 
      Height          =   2055
      Left            =   8160
      Picture         =   "Form5.frx":BF5E
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.PictureBox Picture4 
      Height          =   1935
      Left            =   8160
      Picture         =   "Form5.frx":12236
      ScaleHeight     =   1875
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   240
      Picture         =   "Form5.frx":17D53
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   5280
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   360
      Picture         =   "Form5.frx":1D60D
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   240
      Picture         =   "Form5.frx":234CC
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "FLTRI Road Glide"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "FLHTCUI Ultra Classic Electra Glide"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "FLHTC/FLHTCI Electra Glide Classic"
      Height          =   495
      Left            =   8160
      TabIndex        =   11
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "FLHT/FLHTI Electra Glide Standard"
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "FLHRS/FLHRSI Road King Custom"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "FLHRCI Road King Classic"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "FLHR/FLHRI Road King"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   3255
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form5.Hide
End Sub

