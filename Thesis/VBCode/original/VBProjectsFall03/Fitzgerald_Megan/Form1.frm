VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   660
   ClientTop       =   840
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9120
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   5640
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdOurStory 
      BackColor       =   &H00FF8080&
      Caption         =   "Our Story"
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdMissionTrips 
      BackColor       =   &H00FF8080&
      Caption         =   "Mission Trips"
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdMissionStatement 
      BackColor       =   &H00FF8080&
      Caption         =   "Mission Statement"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Form1.frx":1BAC2
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Serving His Poor..."
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Amigos For Christ"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdMissionStatement_Click()



End Sub
