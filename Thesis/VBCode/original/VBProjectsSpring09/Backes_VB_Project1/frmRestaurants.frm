VERSION 5.00
Begin VB.Form frmRestaurantsPC 
   BackColor       =   &H00800080&
   Caption         =   "Restaurants at Park Central"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click to go back to previous page"
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   4200
      Picture         =   "frmRestaurants.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   120
      Picture         =   "frmRestaurants.frx":4D1A
      ScaleHeight     =   2835
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdlobby 
      BackColor       =   &H0000FF00&
      Caption         =   "Click here to find out more about the lobby lounge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdCafe 
      BackColor       =   &H00FFFF80&
      Caption         =   "Click here to find out more about the Cafe New York "
      BeginProperty Font 
         Name            =   "DotumChe"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   2295
   End
End
Attribute VB_Name = "frmRestaurantsPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form tells the user about two of the Restaurants
'at the Park Central and also tell them what types of food is served
'through a message box
Option Explicit
Private Sub cmdBack_Click()
'allows the user to go back to the room selection page
frmRestaurantsPC.Hide
frmRoomPC.Show

End Sub


Private Sub cmdCafe_Click()
'if the user clicks the Cafe button then this message somes up
MsgBox ("Cafe New York offers a wide varitey of delicious food and is open for breakfast, lunch and dinner")

End Sub

Private Sub cmdlobby_Click()
'if the user choses the lobby then this message comes up
MsgBox ("The lobby lounge offers a wide variety of drinks and your choice of appetizers")
End Sub

Private Sub cmdquit_Click()
End
End Sub
