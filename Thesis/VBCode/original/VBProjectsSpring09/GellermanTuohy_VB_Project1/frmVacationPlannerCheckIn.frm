VERSION 5.00
Begin VB.Form frmVacationPlannerCheckIn 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16125
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   16125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Wonderful Johnnie Travel Experience :'("
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   12960
      TabIndex        =   5
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton cmdVacationSpots 
      BackColor       =   &H000000FF&
      Caption         =   "See All The Wonderful Vaction Spots!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4920
      TabIndex        =   4
      Top             =   7080
      Width           =   6375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Come See The Destinations We Have To Offer For You And Start Planning Your Dream Vacation!!!!!!!!!!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1935
      Left            =   3600
      TabIndex        =   3
      Top             =   4440
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Welcome To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label label 
      BackColor       =   &H00FFFF80&
      Caption         =   "Johnnie Travel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   4560
      TabIndex        =   0
      Top             =   1800
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Height          =   2415
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   7935
   End
End
Attribute VB_Name = "frmVacationPlannerCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Vacation Planner
'Form Name: Destination
'Authors: Luke Gellerman and Tan Tuohy
'3/21/09
'This project will allow someone to arange their vacation, through our travel guide.
'We start by having them choose a destination and whatever destination they choose,
'they will have different options of activities, based on their travel destination.
'Before they choose their activities, they will select a hotel and a flight package.
'They will be taken to a specific Activities page later in the program, based on the Location that is saved on the Destination form.
'The total cost will be projected on the final screen, along with all of their information regarding their vacation.


Option Explicit

Private Sub cmdQuit_Click()
    End 'ends the program
End Sub

Private Sub cmdVacationSpots_Click()
'this button brings the user to the Destination form in the program

    frmVacationPlannerCheckIn.Hide
    frmDestination.Show
    
End Sub


Private Sub Form_Load()
    'This code centers the form on computer screen upon loading

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
