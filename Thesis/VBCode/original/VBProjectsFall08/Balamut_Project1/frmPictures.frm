VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H00000000&
   Caption         =   "Pictures of Weezer"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBlue 
      Caption         =   "Weezer (Blue Album)"
      Height          =   975
      Left            =   1920
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "Weezer (Red Album)"
      Height          =   975
      Left            =   1920
      TabIndex        =   12
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdMB 
      Caption         =   "Make Believe"
      Height          =   975
      Left            =   1920
      TabIndex        =   11
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdMaladroit 
      Caption         =   "Maladroit"
      Height          =   975
      Left            =   1920
      TabIndex        =   10
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdGreen 
      Caption         =   "Weezer (Green Album)"
      Height          =   975
      Left            =   1920
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdPink 
      Caption         =   "Pinkerton"
      Height          =   975
      Left            =   1920
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Go Back to the Info Page"
      Height          =   975
      Left            =   1560
      TabIndex        =   7
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit This Rad Program"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   8400
      Width           =   1095
   End
   Begin VB.PictureBox picPicture 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   3240
      ScaleHeight     =   9075
      ScaleWidth      =   8235
      TabIndex        =   5
      Top             =   240
      Width           =   8295
   End
   Begin VB.CommandButton cmdPat 
      Caption         =   "Pat"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrian 
      Appearance      =   0  'Flat
      Caption         =   "Brian"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdScott 
      Caption         =   "Scott"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdRivers 
      BackColor       =   &H8000000E&
      Caption         =   "Rivers"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Which picture do you want to see?"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Weezer
'Form Name: frmPictures.frm
'Author: Emily Balamut
'Date Written: 10/26/08
'Objective: This form lets the user click on a number of buttons to display certain
'pictures for the buttons.
Option Explicit

Private Sub cmdBlue_Click()
    picPicture.Picture = LoadPicture(App.Path & "\WeezerBlue.jpg")
End Sub

Private Sub cmdBrian_Click()
    picPicture.Picture = LoadPicture(App.Path & "\Brian_Bell_Weezer.jpg")
End Sub

Private Sub cmdGreen_Click()
    picPicture.Picture = LoadPicture(App.Path & "\weezergreen.jpg")
End Sub

Private Sub cmdMain_Click()
    frmPictures.Hide
    frmInfo.Show
End Sub

Private Sub cmdMaladroit_Click()
    picPicture.Picture = LoadPicture(App.Path & "\weezer-maladroit.jpg")
End Sub

Private Sub cmdMB_Click()
    picPicture.Picture = LoadPicture(App.Path & "\weezermakebelieve.jpg")
End Sub

Private Sub cmdPat_Click()
    picPicture.Picture = LoadPicture(App.Path & "\patwilson.jpg")
End Sub

Private Sub cmdPink_Click()
    picPicture.Picture = LoadPicture(App.Path & "\pinkerton.bmp")
End Sub

Private Sub cmdQuit_Click()
MsgBox "Thanks for rocking out with Weezer, " & UserName & "! See you later!", , "Leave"
End
End Sub

Private Sub cmdRed_Click()
    picPicture.Picture = LoadPicture(App.Path & "\redalbumsmall.bmp")
End Sub

Private Sub cmdRivers_Click()
    picPicture.Picture = LoadPicture(App.Path & "\riverscuomo.jpg")
End Sub

Private Sub cmdScott_Click()
    picPicture.Picture = LoadPicture(App.Path & "\scottshriner.jpg")
End Sub
