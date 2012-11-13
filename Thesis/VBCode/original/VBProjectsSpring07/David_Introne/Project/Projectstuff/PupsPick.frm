VERSION 5.00
Begin VB.Form PupsPick 
   BackColor       =   &H00004040&
   Caption         =   "Pick Out A Puppy!"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ba    ck    t    o      Pl ay ers"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Txtplayer 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3960
      TabIndex        =   0
      Text            =   "Pick out a Puppy"
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image Duchsand 
      Height          =   3135
      Left            =   6600
      Picture         =   "PupsPick.frx":0000
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   4050
   End
   Begin VB.Image MtdDog 
      Height          =   3135
      Left            =   480
      Picture         =   "PupsPick.frx":2835A
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Image PitTerrer 
      Height          =   3015
      Left            =   6600
      Picture         =   "PupsPick.frx":4BEBE
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4020
   End
   Begin VB.Image Shepard 
      Height          =   3015
      Left            =   480
      Picture         =   "PupsPick.frx":574F9
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "PupsPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
PupsPick.Hide 'shows puppy profile based on click
Player.Show
End Sub

Private Sub Duchsand_Click()
PupsPick.Hide 'shows puppy profile based on click
Duchpro.Show
puppick = 14
End Sub

Private Sub MtdDog_Click()
puppick = 13
PupsPick.Hide 'shows puppy profile based on click
MtnPro.Show
End Sub

Private Sub PitTerrer_Click()
PupsPick.Hide 'shows puppy profile based on click
PitPro.Show
puppick = 12
End Sub

Private Sub Shepard_Click()
PupsPick.Hide
ShepProp.Show 'shows puppy profile based on click
puppick = 11
End Sub
