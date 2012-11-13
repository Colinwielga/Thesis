VERSION 5.00
Begin VB.Form frmVehicles 
   BackColor       =   &H00404040&
   Caption         =   "Fire Department Vehicles "
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   5760
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   1455
      Left            =   2160
      ScaleHeight     =   1395
      ScaleWidth      =   4155
      TabIndex        =   6
      Top             =   4200
      Width           =   4215
   End
   Begin VB.CommandButton cmdUnit4 
      Caption         =   "Ambulance"
      Height          =   735
      Left            =   6600
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrass5 
      Caption         =   "Grass Rig"
      Height          =   735
      Left            =   6600
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdGator 
      Caption         =   "Gator"
      Height          =   735
      Left            =   6600
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdEngine3 
      Caption         =   "Engine 3"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdLadder2 
      Caption         =   "Ladder 2"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdEngine1 
      Caption         =   "Engine 1 "
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image imgGator 
      Height          =   3855
      Left            =   2040
      Picture         =   "frmVehicles.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image imgGrass5 
      Height          =   3855
      Left            =   2040
      Picture         =   "frmVehicles.frx":7AECB
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image imgUnit4 
      Height          =   3840
      Left            =   2040
      Picture         =   "frmVehicles.frx":F4360
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   4440
   End
   Begin VB.Image imgEngine3 
      Height          =   3840
      Left            =   2040
      Picture         =   "frmVehicles.frx":16CDD4
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   4440
   End
   Begin VB.Image imgEngine1 
      Height          =   3840
      Left            =   2040
      Picture         =   "frmVehicles.frx":1E7F7A
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   4440
   End
   Begin VB.Image imgLadder2 
      Height          =   3840
      Left            =   2040
      Picture         =   "frmVehicles.frx":261D9E
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   4440
   End
End
Attribute VB_Name = "frmVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Saint John's Fire Department
'Form Name: frmVehicles (Vehicle info page)
'Authors: JT Trujillo and Matt Mollet
'Date Written: 2/23/2010
'Objective: To show the user the Fire Department's vehicles, and
            'to inform them of what the vehicles are used for.

'Show engine 1
Private Sub cmdEngine1_Click()

picResults.Cls
imgEngine1.Visible = True
imgUnit4.Visible = False
imgLadder2.Visible = False
imgGrass5.Visible = False
imgEngine3.Visible = False
imgGator.Visible = False
'explain what engine 1 is used for
picResults.Print "Engine 1 is used to carry extrication tools,"
picResults.Print "such as the Jaws of Life, and to connect to"
picResults.Print "the fire hydrant so that the initial team "
picResults.Print "on scene has a water supply to enter the"
picResults.Print "building with."

End Sub

'show engine 3
Private Sub cmdEngine3_Click()

picResults.Cls
imgEngine3.Visible = True
imgUnit4.Visible = False
imgLadder2.Visible = False
imgEngine1.Visible = False
imgGrass5.Visible = False
imgGator.Visible = False
'explain engine 3's uses
picResults.Print "Engine 3 is used for the same things as"
picResults.Print "Engine 1, except it doesn't carry extrication"
picResults.Print "tools.  It is mostly used as a truck to"
picResults.Print "carry a second team of Firefighters to the scene."


End Sub


'show gator
Private Sub cmdGator_Click()

picResults.Cls
imgGator.Visible = True
imgUnit4.Visible = False
imgLadder2.Visible = False
imgEngine1.Visible = False
imgGrass5.Visible = False
imgEngine3.Visible = False
'explain gator's uses
picResults.Print "The Gator is used for carrying extra materials"
picResults.Print "and supplies around to different locations.  It"
picResults.Print "also connects to a small trailer which holds a"
picResults.Print "power generator and a few other items."

End Sub

'show grass rig
Private Sub cmdGrass5_Click()

picResults.Cls
imgGrass5.Visible = True
imgUnit4.Visible = False
imgLadder2.Visible = False
imgEngine1.Visible = False
imgEngine3.Visible = False
imgGator.Visible = False
'explain grass rig's uses
picResults.Print "Grass 5 has a small water tank and hose in the bed"
picResults.Print "with which to fight smaller grass fires.  It also"
picResults.Print "holds wildland firefighting tools and equipment."
picResults.Print "Also, it holds a small boat for water rescue, in the"
picResults.Print "event that someone in one of the surrounding lakes"
picResults.Print "needs assistance for various reasons."

End Sub

'show ladder truck
Private Sub cmdLadder2_Click()

picResults.Cls
imgLadder2.Visible = True
imgUnit4.Visible = False
imgEngine1.Visible = False
imgGrass5.Visible = False
imgEngine3.Visible = False
imgGator.Visible = False
'explain reasons for using ladder truck
picResults.Print "The Ladder Truck is intended for attacking fires"
picResults.Print "from a high angle, it is useful for roof fires or"
picResults.Print "fires in the upper levels of multi-story buildings."
picResults.Print "It is also used to assist people to safety from"
picResults.Print "upper level windows."

End Sub

'return to main form
Private Sub cmdReturn_Click()

frmVehicles.Visible = False
frmMain.Visible = True
End Sub

'show ambulance
Private Sub cmdUnit4_Click()

picResults.Cls
imgUnit4.Visible = True
imgLadder2.Visible = False
imgEngine1.Visible = False
imgGrass5.Visible = False
imgEngine3.Visible = False
imgGator.Visible = False
'explain to user what the ambulance is used for
picResults.Print "The Ambulance, or Unit 4, is used to respond to"
picResults.Print "medical emergencies and it carries a vast array of"
picResults.Print "medical equipment."

End Sub

