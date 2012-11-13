VERSION 5.00
Begin VB.Form frmBegin 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   Picture         =   "NFLdraft.frx":0000
   ScaleHeight     =   5190
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin Draft "
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      Picture         =   "NFLdraft.frx":7D32
      TabIndex        =   0
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2006 NFL Draft"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NFL Draft
'"frmBegin"
'Patrick Hooper
'3/24/06
'the objective of this project is to assist in the decisions of picking a draft pick for the 2006 NFL draft
'it is a shortened and simplified version not containing every available player
Option Explicit
Private Sub cmdBegin_Click()
    frmBegin.Hide
    frmWarRoom.Show
End Sub
