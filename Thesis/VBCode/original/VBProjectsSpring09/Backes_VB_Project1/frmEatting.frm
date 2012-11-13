VERSION 5.00
Begin VB.Form frmRestaurantsMarriott 
   BackColor       =   &H00004080&
   Caption         =   "Restaurants for the Marriott"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000040C0&
      Caption         =   "click to go back to previous page"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   3495
      Left            =   4800
      Picture         =   "frmEatting.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton cmdJW 
      BackColor       =   &H00008080&
      Caption         =   "J.W's Steakhouse"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdbar 
      BackColor       =   &H00FFFF80&
      Caption         =   "Champion's Sports Bar"
      BeginProperty Font 
         Name            =   "Myriad Web Pro Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   120
      Picture         =   "frmEatting.frx":7A8B
      ScaleHeight     =   2595
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmRestaurantsMarriott"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form tells the user about two of the Restaurants
'at the Marriott and also tell them what types of food is served
'through a message box

Option Explicit

Private Sub cmdBack_Click()
'allows the user to go back to the room selection form
frmRestaurantsMarriott.Hide
frmRoomMarriott.Show
End Sub

Private Sub cmdbar_Click()
'if the user clicks on the Champions button then this message comes up
MsgBox ("Champion's Sports Bar serves American Style food and is open for dinner")

End Sub

Private Sub cmdJW_Click()
'if the user clicks on JW then this message comes up
MsgBox ("J.W's Steakhouse specializes in premium steaks and also features fresh seafood and is open for dinner")

End Sub

Private Sub cmdquit_Click()
End
End Sub
