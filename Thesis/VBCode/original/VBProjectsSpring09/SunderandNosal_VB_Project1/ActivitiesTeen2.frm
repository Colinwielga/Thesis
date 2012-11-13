VERSION 5.00
Begin VB.Form frmActivitiesTeen2 
   BackColor       =   &H00C000C0&
   Caption         =   "Alaskan Activities Teen"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   5640
      Picture         =   "ActivitiesTeen2.frx":0000
      ScaleHeight     =   5115
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   1200
      Width           =   4575
   End
   Begin VB.CommandButton cmdReturn3 
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label lbljdhjdf 
      BackColor       =   &H00FFC0FF&
      Caption         =   $"ActivitiesTeen2.frx":74F2
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lblTeenActivities 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Teen Activities"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmActivitiesTeen2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmActivitiesTeen2
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form gives simple information as to what types of activities there are to do on the
'cruise ship that fits people ages 13 to 20, as well as shows a few images of some of those activities.

Private Sub cmdReturn3_Click()
frmActivitiesTeen2.Hide
frmAlaskanHome.Show
End Sub
