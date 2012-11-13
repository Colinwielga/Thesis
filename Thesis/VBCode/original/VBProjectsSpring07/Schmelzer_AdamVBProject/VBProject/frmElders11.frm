VERSION 5.00
Begin VB.Form frmElders11 
   Caption         =   "The Elders"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   4
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdLouis 
      Caption         =   "I am the State"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "How may I prove my cause to the council?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdWait 
      Caption         =   "We shall wait one month to further understand the oracles"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmElders11.frx":0000
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   8640
      Left            =   0
      Picture         =   "frmElders11.frx":00C0
      Top             =   0
      Width           =   10995
   End
End
Attribute VB_Name = "frmElders11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'to give the user different options through command buttons
'this form gives the user options on how to deal with the elders, which mainly affects
'the elderpoints variable but also resources if certain options are taken as time is
'considered to be a resource

Private Sub cmdLouis_Click()
Resources = Resources - 0
Elderpoints = Elderpoints - 2
frmElders11.Hide
frmPeople1.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdTest_Click()
frmElders11.Hide
frmEldertest.Show
End Sub

Private Sub cmdWait_Click()
Resources = Resources - 100
Elderpoints = Elderpoints + 1
frmElders11.Hide
frmPeople1.Show
End Sub
