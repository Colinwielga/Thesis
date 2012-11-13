VERSION 5.00
Begin VB.Form frmRoomNumber 
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   Picture         =   "RoomNumber.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000080FF&
      Caption         =   "Return to Information"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H000080FF&
      Caption         =   "Insert your key to see your room information"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   5655
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   600
      ScaleHeight     =   2355
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
   End
   Begin VB.CommandButton cmdCheckOut 
      BackColor       =   &H000080FF&
      Caption         =   "Proceed to Check Out"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   2895
   End
End
Attribute VB_Name = "frmRoomNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hotel Check In
'frmRoomNumber
'Shannon Hooley
'10/16/09
'This form allows the guest to see a recap of what their room is, and tell them how to get a hold of the front desk

Private Sub cmdCheckOut_Click()
frmRoomNumber.Hide
frmCheckOut.Show
End Sub

Private Sub cmdKey_Click()
'clears the past guest's room info
picResults.Cls
'prints the current guest's info
picResults.Print "You are in room 101"
picResults.Print "  If we can do anything to make your stay more comfortable "
picResults.Print "  please don 't hesitate to press '0' on your phone"
picResults.Print "      Press 'Check Out' when you are ready to leave"
End Sub

Private Sub cmdReturn_Click()
frmRoomNumber.Hide
frmInfo.Show
End Sub
