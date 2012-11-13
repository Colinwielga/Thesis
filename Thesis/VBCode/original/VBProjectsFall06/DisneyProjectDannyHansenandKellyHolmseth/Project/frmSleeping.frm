VERSION 5.00
Begin VB.Form frmSleeping 
   BackColor       =   &H00FF0000&
   Caption         =   "Sleeping Beauty"
   ClientHeight    =   8265
   ClientLeft      =   2925
   ClientTop       =   1500
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   Picture         =   "frmSleeping.frx":0000
   ScaleHeight     =   8265
   ScaleWidth      =   9675
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   5640
      Picture         =   "frmSleeping.frx":305A
      ScaleHeight     =   2475
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   5640
      Width           =   3735
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000C0C0&
      Caption         =   "Back"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   $"frmSleeping.frx":19FD8
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSleeping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form  is to display to the user a summary of the movie "Sleeping Beauty"
Private Sub cmdBack_Click()
frmSleeping.Hide
frmTop.Show
End Sub

