VERSION 5.00
Begin VB.Form frmCheckIn 
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   Picture         =   "CheckIn.frx":0000
   ScaleHeight     =   8580
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H00FFFF00&
      Caption         =   "You are checking prices"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFF00&
      Caption         =   "Return to Check In"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton cmdUpgrade 
      BackColor       =   &H00FFFF00&
      Caption         =   "See what the rooms include"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4440
      MaskColor       =   &H00000040&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdWalkIn 
      BackColor       =   &H00FFFF00&
      Caption         =   "See what rooms are avaliable"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdReservation 
      BackColor       =   &H00FFFF00&
      Caption         =   "You had a reservation"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hotel Check In
'frmCheckIn
'Shannon Hooley
'10/16/09
'This form is the main hub for the program
'From here you get get to the description of what each room offers, see what rooms are avaliable
'the guest can name their price, as well as head to the check in area

Private Sub cmdPrice_Click()
Dim Prices As Long
'lets the guest name their own price
Prices = InputBox("Please enter the amount you wish to pay for a room", "Prices")
'tells the guest what that amount can buy them
Select Case Prices
Case Is >= 453
    MsgBox ("You can get a Presidential Suite for only $453 a night.")
Case Is >= 379
    MsgBox ("You can get a Suite for only $379 a night.")
Case Is >= 206
    MsgBox ("You can get a King sized room for only $206 a night.")
Case Is >= 150
    MsgBox ("You can get a Queen sized room for only $150 a night.")
Case Is >= 109
    MsgBox ("You can get a Double bed room for only $109 a night.")
Case Else
    MsgBox ("Unfortunatly we don't have rooms for that price.")
End Select

End Sub

Private Sub cmdReservation_Click()
'brings the guest to the information sheet
frmCheckIn.Hide
frmInfo.Show
End Sub

Private Sub cmdReturn_Click()
'brings the guest back to the hotel lobby
frmCheckIn.Hide
frmHotelLobby.Show
End Sub

Private Sub cmdUpgrade_Click()
'brings the guest to the sheet that explains what each room offers
frmCheckIn.Hide
frmRooms.Show
End Sub

Private Sub cmdWalkIn_Click()
'allows the guest to see what rooms are availiable and where that room is located
frmCheckIn.Hide
frmLayout.Show
End Sub
