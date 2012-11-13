VERSION 5.00
Begin VB.Form frmRestaurantsHilton 
   BackColor       =   &H008080FF&
   Caption         =   "Places to eat at the Hilton "
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0000FFFF&
      Caption         =   "quit"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FF80&
      Caption         =   "click to go back to previous page"
      Height          =   735
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   5040
      Picture         =   "frmFood.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   120
      Picture         =   "frmFood.frx":2648
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CommandButton cmdAndiamo 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to check out Andiamo"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdFood 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click here to find out more about The Cafe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmRestaurantsHilton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form tells the user about the Restaurants at the
'Hilton and the food secletion pops up through a message box

Option Explicit

Private Sub cmdAndiamo_Click()
'if the user clicks on Andiamo then this message appears
MsgBox ("Andiamo is a italian restaurant that offers wonderful italian food and is open for Dinner")

End Sub

Private Sub cmdBack_Click()
'allows the user to go back to the room selection page
frmRestaurantsHilton.Hide
frmRoomsHilton.Show

End Sub

Private Sub cmdFood_Click()
'if the user clicks on the Cafe then this message comes up
MsgBox ("The Cafe is open for Breakfast, Lunch and Dinner and offers a wide variety of international cuisine")
End Sub


Private Sub cmdquit_Click()
End
End Sub
