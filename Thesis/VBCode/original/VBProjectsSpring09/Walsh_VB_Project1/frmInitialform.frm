VERSION 5.00
Begin VB.Form frmInitialform 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main Menu"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   Picture         =   "frmInitialform.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdOwn 
      BackColor       =   &H8000000A&
      Caption         =   "Figure out my own Statistics"
      Height          =   1095
      Left            =   6000
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdTwins 
      Caption         =   "2008 Twins Stats"
      Height          =   1095
      Left            =   2160
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Would you like to view the 2008 Minnesota Twins Baseball Statistics or figure out statistics of your own!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   7335
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to My Baseball Statistics Program!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "frmInitialform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Baseball Batting Statistics
'frmInitialform
'Aaron Walsh
'March 24, 2009
'This project will figure out various batting statistics like BA, OPS, OBP, and SLG
'by inputting either a file of twins players or user generated numbers for
'certain batting catagories

Private Sub cmdOwn_Click()
    frmInitialform.Hide
    frmOwnInput.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdTwins_Click()
    frmInitialform.Hide
    frmReadTwinsStats.Show
End Sub


