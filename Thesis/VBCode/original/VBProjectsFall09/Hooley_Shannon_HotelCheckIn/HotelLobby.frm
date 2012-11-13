VERSION 5.00
Begin VB.Form frmHotelLobby 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   Picture         =   "HotelLobby.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLeave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leave the Lake Front Inn"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8040
      Width           =   3495
   End
   Begin VB.CommandButton cmdCheckIn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Check In"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   3615
   End
   Begin VB.Label lblGlad 
      BackColor       =   &H00400000&
      Caption         =   "We are glad you chose to stay with us."
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00400000&
      Caption         =   "Welcome to the Lake Front Inn"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   5175
   End
End
Attribute VB_Name = "frmHotelLobby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hotel Check In
'frmHotelLobby
'Shannon Hooley
'10/16/09
'This project is to help people check into a hotel room, let them explore the different types of rooms
'Allow them to name a price and see what they can get for that amount, enter in their information
'This program also allows the gues sto check out and see their bill
'This formis bringing the guest from the "hotel lobby" to the "desk" to explore what the hotel has to offer

Private Sub Picture1_Click()

End Sub

Private Sub cmdCheckIn_Click()
'this brings the guest from the "hotel lobby" to the "front desk"
frmHotelLobby.Hide
frmCheckIn.Show

End Sub

Private Sub cmdLeave_Click()
End
End Sub
