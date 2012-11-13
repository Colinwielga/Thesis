VERSION 5.00
Begin VB.Form frmHome 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcitations 
      Caption         =   "work cited"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "View Wrestlers in Each Weight Class"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   5
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton cmdShop 
      Caption         =   "Get Your Johnnie Gear"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdCompetition 
      Caption         =   "Become a Johnnie Wrestler"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmdRoster 
      Caption         =   "View Current Johnnie Wrestlers"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton cmdSchedule 
      Caption         =   "View 2009-2010 Schedule"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Picture         =   "frmHome.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   6765
      Left            =   0
      Picture         =   "frmHome.frx":6E206
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcitations_Click()
frmHome.Hide 'hide the home page
frmcited.Show 'show the work cited page
End Sub

Private Sub cmdCompetition_Click()
frmHome.Hide 'hides the home page
frmCompetition.Show 'shows the Competition form

End Sub

Private Sub cmdQuit_Click()
    End ' ends the program
End Sub

Private Sub cmdRoster_Click()
frmHome.Hide 'hides the home page
frmRoster.Show 'shows the roster

End Sub

Private Sub cmdSchedule_Click()
frmHome.Hide 'hides the home page
frmSchedule.Show 'shows the schedule
End Sub

Private Sub cmdSearch_Click()
frmSearch.Show 'shows the search page
frmHome.Hide 'hides the home page
End Sub

Private Sub cmdShop_Click()
frmshopfinal.Show 'shows the shop
frmHome.Hide ' hides the home page
End Sub
