VERSION 5.00
Begin VB.Form frmBoat 
   BackColor       =   &H00C0C0C0&
   Caption         =   "frmBoat"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   Picture         =   "frmQuiz.frx":0000
   ScaleHeight     =   9825
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00000080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdRower 
      BackColor       =   &H00000080&
      Caption         =   "Push Me!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdphrase 
      BackColor       =   &H00000080&
      Caption         =   "Learn Some Crew Phrases!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdCox 
      BackColor       =   &H00000080&
      Caption         =   "Push Me!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdPort 
      BackColor       =   &H00000080&
      Caption         =   "Push Me!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdStar 
      BackColor       =   &H00000080&
      Caption         =   "Push Me!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdBow 
      BackColor       =   &H00000080&
      Caption         =   "Push Me!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton cmdStern 
      BackColor       =   &H00000080&
      Caption         =   "Push Me!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00000080&
      Caption         =   "Return to the Main Screen"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   2295
   End
   Begin VB.Label lblPush 
      Alignment       =   2  'Center
      Caption         =   "Push a Button to Learn More Information About the Boats the Crew Team Uses and Some Phrases"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3375
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblBoat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Boat Basics"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1695
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "frmBoat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: CSB/SJU Crew
'Form name: frmMeettheMembers
'Authors: Lauren Nephew and Rachel Stalley
'Date: October 16th, 2009
'Objective: To have the user click on a button and see a message box pop up with information on the area of the crew boat they clicked on.
Option Explicit

Private Sub cmdBow_Click() 'This button describes the bow
    MsgBox "This is called the bow. It is the back of the boat. The rowers have their backs toward this end.", , "Information"
End Sub

Private Sub cmdCox_Click() 'This button describes the coxswain
    MsgBox "The person who sits here is called the coxswain or the cox. They do not row, instead they tell the rowers what to do, keep them together, and keep their direction.", , "Information"
End Sub

Private Sub cmdphrase_Click() 'This button gives common phrases
    MsgBox "Some common phrases include: Wain off -- meaning to stop, Cox Box -- the microphone the cox uses in the boat, Up and Over Head -- lifting the boat out of the water and each rower holding the boat above their head to transport the boat."
           
End Sub

Private Sub cmdPort_Click() 'This button describes port
    MsgBox "This is the port side. It is the rowers's right side and the cox's left. Half of the oars are on this side of the boat.", , "Information"
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
'This brings the user back to the main menu screen
frmCSBSJUCrewMain.Show
frmBoat.Hide
End Sub

Private Sub cmdRower_Click() 'This button gives information on the rower
    MsgBox "This is one of 8, 4, or 2 rowers in a boat. They hold one oar either on the star side or the port side. They move backward on the water and follow the guidance of the cox. They put their feet into shoes that are strapped down, and sit on seats that slide back and forth with their stroke.", , "Information"
End Sub

Private Sub cmdStar_Click() 'This button describes star
    MsgBox "This is the star side. It is the rowers' left side and the cox's right. Half of the oars are on this side of the boat.", , "Information"
End Sub

Private Sub cmdStern_Click() 'This button describes the stern
    MsgBox "This is called the stern. It is the front of the boat. All of the rowers face toward the stern.", , "Information"
End Sub
