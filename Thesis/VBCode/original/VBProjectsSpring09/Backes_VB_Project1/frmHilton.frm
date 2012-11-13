VERSION 5.00
Begin VB.Form frmHilton 
   BackColor       =   &H00400040&
   Caption         =   "Hilton"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdroom 
      BackColor       =   &H0000FF00&
      Caption         =   "click here to see our room options and figure our your room total"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H008080FF&
      Caption         =   "Click to go back to start page"
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   4560
      Picture         =   "frmHilton.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblHilton 
      BackColor       =   &H00FFFF80&
      Caption         =   "Thanks for picking the Hilton we are happy that you will be spending your vacation with us!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmHilton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form is the welcome page for the Hilton hotel
'The user can click on it to find out about rooms, restaurants and
'activities

Option Explicit



Private Sub cmdBack_Click()
'allows the uset to go back to the starting page
frmHilton.Hide
frmOpen.Show

End Sub


Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdroom_Click()
'allows the uset to go to the room selection page
frmHilton.Hide
frmRoomsHilton.Show

End Sub
