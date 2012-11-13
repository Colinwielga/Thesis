VERSION 5.00
Begin VB.Form frmTH 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
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
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   6600
      Width           =   2535
   End
   Begin VB.PictureBox picBoxTH 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   240
      Picture         =   "frmTH.frx":0000
      ScaleHeight     =   4335
      ScaleWidth      =   5295
      TabIndex        =   2
      Top             =   1920
      Width           =   5295
   End
   Begin VB.PictureBox picBoxSJU 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmTH.frx":575F
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
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblTrue 
      BackColor       =   &H00000000&
      Caption         =   "True Story: When Tim is not on the Lacrosse field, he still uses his stick to keep his hordes of adoring female fans at bay."
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
      Height          =   2415
      Left            =   6120
      TabIndex        =   5
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lblLoves 
      BackColor       =   &H00000000&
      Caption         =   "Loves: The Ladies  Does Not Love: People who are jealous of his skills"
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
      Height          =   1695
      Left            =   6120
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblTH 
      BackColor       =   &H00000000&
      Caption         =   "Tim ""Ent"" Herby"
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
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
End
Attribute VB_Name = "frmTH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmTH (frmTH.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Tim Herby that can be linked to the team bio page

Private Sub cmdBack_Click()
    frmTH.Visible = False
    frmBio.Visible = True
End Sub

