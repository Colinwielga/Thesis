VERSION 5.00
Begin VB.Form frmRide 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   2505
   ClientTop       =   1320
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   10590
   Begin VB.TextBox txtSelectType 
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   21
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4320
      TabIndex        =   12
      ToolTipText     =   "Enter either Trail, Touring, Mountain, or Performance"
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H8000000E&
      Caption         =   "Help!"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSpeed 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Calculate Engine Size"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtSpeed 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Input desired speed to match engine size"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox picOutput 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      ScaleHeight     =   675
      ScaleWidth      =   4755
      TabIndex        =   7
      Top             =   4440
      Width           =   4815
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H80000003&
      Caption         =   "I have calculated my desired speed and entered my snowmobile type above! Click here to move on"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000E&
      Caption         =   "Go back to previous screen"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chris Donnelly"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   14
      Top             =   8520
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Now enter your desired snowmobile type below"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   13
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Enter a desired to speed (in MPH) to find an estimate engine size in cubic centimeters (C.C.'s)"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   10
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Touring"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "Performance"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "Mountain"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Trail"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image imgPerformance 
      Height          =   1890
      Left            =   7320
      Picture         =   "frmRide.frx":0000
      Top             =   4200
      Width           =   3240
   End
   Begin VB.Image imgTouring 
      Height          =   2430
      Left            =   7320
      Picture         =   "frmRide.frx":8C0C
      Top             =   1800
      Width           =   3240
   End
   Begin VB.Image imgMountain 
      Height          =   2115
      Left            =   120
      Picture         =   "frmRide.frx":F7E6
      Top             =   3600
      Width           =   3240
   End
   Begin VB.Image imgTrail 
      Height          =   1875
      Left            =   120
      Picture         =   "frmRide.frx":16392
      Top             =   1560
      Width           =   3240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   $"frmRide.frx":1FBC2
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "frmRide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InputBox As Integer
'will go back to main menu
Private Sub cmdBack_Click()
frmRide.Hide
frmMain.Show
End Sub
'will display a message
Private Sub cmdHelp_Click()
MsgBox "This screen will help to determine the basic outline of your dream snowmobile. There are many to choose from so use this program to learn more about the selections out there. Enter your maximum desired speed in the center text box, and your desired snowmobile type in the lower center text box and then click on the bottom center button when you're ready. Click on go back to the previous screen to go back at anytime", , "Help"
End Sub
'this subroutine will match the user's maximum desired speed with an engine size
Private Sub cmdSpeed_Click()
picOutput.Cls
      Speed = txtSpeed.Text
        Select Case Speed
        'will classify input into one of the cases and print a message
        'also sets the Speed variable to a specific engine size which will be used
        'on the next form
        Case 0 To 50
            Speed = 380
            Speed1 = 340
            picOutput.Print "380 C.C. engine or less will do just fine"
        Case 51 To 70
            Speed = 500
            Speed1 = 440
            Speed2 = 550
            picOutput.Print "Go for an engine size around 500 C.C."
        Case 71 To 90
            Speed = 600
            Speed1 = 650
            picOutput.Print "A 600 C.C. engine will get you there"
        Case 91 To 105
            Speed = 700
            picOutput.Print "A 700 C.C. engine will easily do that"
        Case 106 To 125
            Speed = 800
            Speed1 = 900
            Speed2 = 1000
            picOutput.Print "If you're this crazy, go for 800 C.C. or more"
        Case Else
            picOutput.Print "Who are you trying to outrun? Enter a smaller number."
    End Select
End Sub
'goes to next form
Private Sub cmdNext_Click()
selectType = txtSelectType.Text
frmRide.Hide
frmSnowmobiles.Show
End Sub
'clicking on this picture will display a message
Private Sub imgTrail_Click()
    MsgBox " Trail snowmobiles offer a standard ride with excellent cornering that hugs the trails common throughout the midwest", , "Trail Snowmobiles"
End Sub
'clicking on this picture will display a message
Private Sub imgMountain_Click()
    MsgBox "Mountain sleds come standard with longer tracks and deeper lugs on the track which keeps you moving in deep, powerdy snow", , "Mountain Sleds"
End Sub
'clicking on this picture will display a message
Private Sub imgTouring_Click()
    MsgBox "Touring sleds offer the deluxe way to snowmobile with passengers in mind. Cruise the winter landscape in comfort as electric start, reverse, driver and passenger handwarmers, and high windshield come standard.", , "Touring Snowmobiles"
End Sub
'clicking on this picture will display a message
Private Sub imgPerformance_Click()
    MsgBox "Hold on and don't blink! Performance snowmobiles pack the most powerful engines and designed for the speed freak.", , "Performance Snowmobiles"
End Sub
