VERSION 5.00
Begin VB.Form frmActivities 
   BackColor       =   &H0000FF00&
   Caption         =   "Activities"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   600
      Picture         =   "Activities.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CommandButton cmdGoBackHome 
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label lblInfo2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Indoor activities include: movie theater, arcades, Playland, Lego room, craft area, and much more!"
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
      Left            =   5400
      TabIndex        =   2
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label lblInfo1 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"Activities.frx":C5B4
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5400
      TabIndex        =   1
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label lblActivities 
      BackColor       =   &H00FFFFC0&
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
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "frmActivities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmActivities
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form gives simple information as to what types of activities there are to do on the
'cruise ship that fits people ages 4 to 12, as well as shows a few images of some of those activities.

Private Sub cmdGoBackHome_Click()
frmActivities.Hide
frmCaribbeanHome.Show
End Sub

