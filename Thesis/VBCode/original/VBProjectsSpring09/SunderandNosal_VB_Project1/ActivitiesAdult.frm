VERSION 5.00
Begin VB.Form frmActivitiesAdult 
   BackColor       =   &H00FF00FF&
   Caption         =   "Activities Adult"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   5040
      Picture         =   "ActivitiesAdult.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   1680
      Width           =   4335
   End
   Begin VB.CommandButton cmdReturn2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Caribbean Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label lblInfo3 
      BackColor       =   &H00FFC0FF&
      Caption         =   $"ActivitiesAdult.frx":50DF
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label lblActivitiesAdult 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Activities Adult"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmActivitiesAdult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmActivitiesAdult
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form gives simple information as to what types of activities there are to do on the
'cruise ship that fits people ages 21 and older, as well as shows a few images of some of those activities.
Private Sub cmdReturn2_Click()
frmCaribbeanHome.Show
frmActivitiesAdult.Hide
End Sub
