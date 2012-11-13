VERSION 5.00
Begin VB.Form frmActivities2 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Alaskan Activities"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   4920
      Picture         =   "Activities2.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   1800
      Width           =   6015
   End
   Begin VB.CommandButton cmdReturn2 
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
      Height          =   1455
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label lblkgkg 
      BackColor       =   &H80000009&
      Caption         =   "Indoors: Penguin Paradise waterpark, movie theater, Play Place, Sundae Ally, arcade, board games, and much more!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   2
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Label lbl3 
      BackColor       =   &H80000009&
      Caption         =   "Outdoors: whale watching from the cruise deck, ice skating rink, igloo playground, snow painting, and so much more!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label lblAlaskanActivities 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Activities for Kids"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "frmActivities2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmActivities2
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form gives simple information as to what types of activities there are to do on the
'cruise ship that fits people ages 4 to 12, as well as shows a few images of some of those activities.

Private Sub cmdReturn2_Click()
frmActivities2.Hide
frmAlaskanHome.Show
End Sub
