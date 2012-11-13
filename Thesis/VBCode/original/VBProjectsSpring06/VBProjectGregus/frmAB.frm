VERSION 5.00
Begin VB.Form frmAB 
   BackColor       =   &H80000007&
   Caption         =   "Player Profile"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Bio Page"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   6960
      Width           =   2535
   End
   Begin VB.PictureBox picBoxAB 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   480
      Picture         =   "frmAB.frx":0000
      ScaleHeight     =   4455
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmAB.frx":4415
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label lblCredit 
      BackColor       =   &H00000000&
      Caption         =   "Project by: Dan Gregus"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblDream 
      BackColor       =   &H00000000&
      Caption         =   "Lifelong Dream: Win World's Strongest Man competition"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4080
      TabIndex        =   6
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label lblLikes 
      BackColor       =   &H00000000&
      Caption         =   "Likes: Solid defense       Dislikes: Ballett, Opera, ""Girly Men"""
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4080
      TabIndex        =   5
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblAB 
      BackColor       =   &H00000000&
      Caption         =   "Adam Benny"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblHobbies 
      BackColor       =   &H80000012&
      Caption         =   "Hobbies: Ultimate Fighting, Sky Diving, Bull Riding"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   4575
   End
End
Attribute VB_Name = "frmAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmAB (frmAB.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Adam Benny that can be linked to the team bio page



'This button brings the user back to the team biography page
Private Sub cmdBack_Click()
    frmAB.Visible = False
    frmBio.Visible = True
End Sub


