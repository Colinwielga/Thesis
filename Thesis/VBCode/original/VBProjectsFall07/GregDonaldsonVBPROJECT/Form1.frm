VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMeetTeam 
      Caption         =   "Meet The Team"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox picsju 
      Height          =   855
      Left            =   3240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearchSwimmer 
      Caption         =   "Search Swimmers"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdBestTimes 
      Caption         =   "Best Times"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblsju 
      Caption         =   "Saint John's University Swimming"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

