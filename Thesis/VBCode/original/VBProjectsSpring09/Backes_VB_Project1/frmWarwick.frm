VERSION 5.00
Begin VB.Form frmWarwick 
   BackColor       =   &H00400000&
   Caption         =   "Warwick Hotel"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF80&
      Caption         =   "click here to go back to the opening page"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdroom 
      BackColor       =   &H0000FF00&
      Caption         =   "Click here to see the room options at the Warwick"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   4335
      Left            =   4320
      Picture         =   "frmWarwick.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Label lblwarwick 
      BackColor       =   &H00C000C0&
      Caption         =   $"frmWarwick.frx":7528
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmWarwick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form is the welcome page for the Warwick hotel
'The user can click on it to find out about rooms, restaurants and
'activities
Option Explicit


Private Sub cmdBack_Click()
'allows the user to go back to the opening form
frmWarwick.Hide
frmOpen.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdroom_Click()
'allows the user to go to the room selection form
frmWarwick.Hide
frmRoomsWarwick.Show
End Sub
