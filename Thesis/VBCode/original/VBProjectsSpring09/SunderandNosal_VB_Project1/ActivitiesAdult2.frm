VERSION 5.00
Begin VB.Form frmActivitiesAdult2 
   BackColor       =   &H000000C0&
   Caption         =   "Alaskan Activities Adult"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   4440
      Picture         =   "ActivitiesAdult2.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   2520
      Width           =   5775
   End
   Begin VB.CommandButton cmdReturn6 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Alaskan Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lbljkjdf 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ActivitiesAdult2.frx":13748
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label lblhjksh 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ActivitiesAdult2.frx":138C3
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adult Activities"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmActivitiesAdult2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmActivitiesAdult2
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form gives simple information as to what types of activities there are to do on the
'cruise ship that fits people ages 21 and older, as well as shows a few images of some of those activities.

Private Sub cmdReturn6_Click()
frmAlaskanHome.Show
frmActivitiesAdult2.Hide
End Sub
