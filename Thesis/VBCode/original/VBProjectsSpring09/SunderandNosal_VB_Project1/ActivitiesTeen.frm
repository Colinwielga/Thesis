VERSION 5.00
Begin VB.Form frmActivitiesTeen 
   BackColor       =   &H00FF8080&
   Caption         =   "Activities Teen"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   5640
      Picture         =   "ActivitiesTeen.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   1560
      Width           =   4695
   End
   Begin VB.CommandButton cmdReturn 
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
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label lblActivitiesTeen 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"ActivitiesTeen.frx":BBF2
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
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label lblActivitiesTeen 
      BackColor       =   &H00FFC0C0&
      Caption         =   "   Activities Teen"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmActivitiesTeen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmActivitiesTeen
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form gives simple information as to what types of activities there are to do on the
'cruise ship that fits people ages 13 to 20, as well as shows a few images of some of those activities.
Private Sub cmdReturn_Click()
frmCaribbeanHome.Show
frmActivitiesTeen.Hide
End Sub

