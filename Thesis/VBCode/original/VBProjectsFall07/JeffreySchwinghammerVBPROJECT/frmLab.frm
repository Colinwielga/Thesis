VERSION 5.00
Begin VB.Form frmLab 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   1725
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   11100
   Begin VB.PictureBox picTure 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   480
      Picture         =   "frmLab.frx":0000
      ScaleHeight     =   6015
      ScaleWidth      =   10455
      TabIndex        =   8
      Top             =   120
      Width           =   10455
   End
   Begin VB.CommandButton cmdNext2 
      Caption         =   "Next"
      Height          =   615
      Left            =   9360
      TabIndex        =   7
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   9240
      TabIndex        =   6
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Previous Room"
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdDesk 
      Caption         =   "Check the Computer Desk"
      Height          =   615
      Left            =   6360
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheckDoor 
      Caption         =   "Check the Far Door"
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdTestTube 
      Caption         =   "Check Test Tubes"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdComputer 
      Caption         =   "Turn the Computer On"
      Height          =   615
      Left            =   9240
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
   Begin VB.PictureBox picLabTxt 
      Height          =   1095
      Left            =   1200
      ScaleHeight     =   1035
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   7320
      Width           =   7695
   End
End
Attribute VB_Name = "frmLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim deskcheck As Boolean
Dim tubecheck As Boolean
Dim doorcheck As Boolean
Dim returncheck As Boolean
Private Sub cmdCheckDoor_Click()
'Hide other commands
cmdTestTube.Enabled = False
cmdReturn.Enabled = False
cmdDesk.Enabled = False
cmdCheckDoor.Enabled = False

picLabTxt.Cls

picLabTxt.Print "You walk up to the far door. It is another reinforced steel door to keep people"
picLabTxt.Print " out... or maybe to keep something in. The sign on the door says 'TESTING"
picLabTxt.Print " AREA' You look through the small viewing window and see nothing but darkness."
cmdNext.Visible = True

If deskcheck = True And tubecheck = True And doorcheck = True And returncheck = True Then
    cmdComputer.Visible = True
End If
End Sub

Private Sub cmdComputer_Click()
    
    MsgBox ("You turn the computer on...")
    frmComputer.Show
    frmLab.Hide
    
End Sub

Private Sub cmdDesk_Click()
picLabTxt.Cls
picLabTxt.Print "You look at the computer desk, it is covered with paperand the computer "
picLabTxt.Print "chair has been knocked over. Blood has been splattered over much of it "
picLabTxt.Print "but has but has dried and begone to spell. Gross."

deskcheck = True
If deskcheck = True And tubecheck = True And doorcheck = True And returncheck = True Then
    cmdComputer.Visible = True
End If
End Sub

Private Sub cmdNext_Click()
picLabTxt.Cls
picLabTxt.Print "Nothing... Is this an empty room...?"
picLabTxt.Print "Then you eye catches something! A red flicker...Eyes?"
picLabTxt.Print "Then you here something scrambling. It runs wildly at the door!"
cmdNext.Visible = False
cmdNext2.Visible = True
End Sub

Private Sub cmdNext2_Click()
picLabTxt.Cls
picLabTxt.Print "It smashes into the steel door with immense force. The sound of the"
picLabTxt.Print "collision is deafening. You jump back expecting it to knock the"
picLabTxt.Print "the door down. Fortunately, it holds strong, keeping in the creature."
cmdNext.Visible = False
doorcheck = True
If deskcheck = True And tubecheck = True And doorcheck = True And returncheck = True Then
    cmdComputer.Visible = True
End If
cmdNext2.Visible = False
cmdTestTube.Enabled = True
cmdReturn.Enabled = True
cmdDesk.Enabled = True
cmdCheckDoor.Enabled = True

End Sub

Private Sub cmdReturn_Click()
picLabTxt.Cls
picLabTxt.Print "You check the door, it is firmly locked. There is no way"
picLabTxt.Print "you could be able to knocked down this reinforced steel door."
returncheck = True

If deskcheck = True And tubecheck = True And doorcheck = True And returncheck = True Then
    cmdComputer.Visible = True
End If

End Sub

Private Sub cmdTestTube_Click()
picLabTxt.Cls

picLabTxt.Print "Several test tubes are lined up across the counter and in the"
picLabTxt.Print "cabinet. You wonder what these are used for... are they dangerous?"
tubecheck = True
If deskcheck = True And tubecheck = True And doorcheck = True And returncheck = True Then
    cmdComputer.Visible = True
End If

End Sub

Private Sub Form_activate()
cmdComputer.Visible = False
cmdNext.Visible = False
cmdNext2.Visible = False
deskcheck = False
tubecheck = False
doorcheck = False
returncheck = False

If labcheck = True Then
    cmdDesk.Visible = False
    cmdCheckDoor.Visible = False
    cmdTestTube.Visible = False
End If
picLabTxt.Cls

picLabTxt.Print "You step through an sliding automatic door into a laboratory."
picLabTxt.Print "The door slides shut behind with a click: an automatic locking system."
picLabTxt.Print "You hear a growling sound... you might not be alone..."

End Sub
