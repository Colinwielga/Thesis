VERSION 5.00
Begin VB.Form frmPanamaCity 
   BackColor       =   &H008080FF&
   Caption         =   "Form4"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form4"
   ScaleHeight     =   8025
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit2 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton cmdDaysInn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Days INN Anyone?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdBeach 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Beach View Hotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label lblwhat 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Where Would You Like To Stay?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   3795
      Left            =   5040
      Picture         =   "frmPanamaCity.frx":0000
      Top             =   2400
      Width           =   4665
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   360
      Picture         =   "frmPanamaCity.frx":39D4A
      Top             =   2400
      Width           =   4500
   End
   Begin VB.Label lblPanamaCity 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Did Someone Say Panama City?"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "frmPanamaCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'creating an otption for people on where they would like to stay'
'Moves them between one frame to the next'
'October 14th, 2009'
'Blake Bauer'
'panama city ro hotel'
Private Sub cmdBeach_Click()
    frmPanamaCity.Hide
    frmHotel.Show
    
End Sub
'panama city ro hotel'
Private Sub cmdDaysInn_Click()
    frmPanamaCity.Hide
    frmHotel.Show
End Sub
'Quit Button'
Private Sub cmdQuit2_Click()
    End
End Sub
