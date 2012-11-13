VERSION 5.00
Begin VB.Form RachelHaney3 
   BackColor       =   &H0000FFFF&
   Caption         =   "RachelHaney3"
   ClientHeight    =   5550
   ClientLeft      =   2955
   ClientTop       =   2265
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7155
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   4755
      TabIndex        =   12
      Top             =   1680
      Width           =   4815
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   975
      Left            =   5520
      Picture         =   "RachelHaneyVBProject3.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   840
      Picture         =   "RachelHaneyVBProject3.frx":4442
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   615
      Left            =   3840
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdBus 
      Caption         =   "Bus"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCoach 
      Caption         =   "Coach"
      Height          =   735
      Left            =   4560
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdBusiness 
      Caption         =   "Business Class"
      Height          =   735
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First Class"
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   5640
      Picture         =   "RachelHaneyVBProject3.frx":7344
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   11
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCar 
      Caption         =   "Car"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblPlane 
      BackColor       =   &H0000FFFF&
      Caption         =   "Plane"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblTransportation 
      BackColor       =   &H00FF80FF&
      Caption         =   "What type of transportation would you like to take?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "RachelHaney3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RachelHaney3 (RachelHaneyVBProject2.frm)
'Rachel Haney 3/11/04
'This form asks people what type of transportation
'they would like to take on their vacation.

Private Sub cmdBus_Click()
    Travel = 5
    Total = Total + 100
    picResults.Print
    picResults.Print "You decided to ride on a bus to your destination."
    cmdContinue.Visible = True
    cmdFirst.Visible = False
    cmdBusiness.Visible = False
    cmdCoach.Visible = False
    cmdCar.Visible = False
    cmdBus.Visible = False
End Sub

Private Sub cmdBusiness_Click()
    Travel = 2
    Total = Total + (150 * People)
    picResults.Print
    picResults.Print "You chose to fly business class to your destination."
    cmdContinue.Visible = True
    cmdFirst.Visible = False
    cmdCoach.Visible = False
    cmdCar.Visible = False
    cmdBus.Visible = False
    cmdBusiness.Visible = False
End Sub

Private Sub cmdCar_Click()
    Travel = 4
    Total = Total + 200
    picResults.Print
    picResults.Print "You decided to drive to your destination."
    cmdContinue.Visible = True
    cmdFirst.Visible = False
    cmdBusiness.Visible = False
    cmdCoach.Visible = False
    cmdBus.Visible = False
    cmdCar.Visible = False
End Sub

Private Sub cmdCoach_Click()
    Travel = 3
    Total = Total + (75 * People)
    picResults.Print
    picResults.Print "You decided to fly coach to your destination."
    cmdContinue.Visible = True
    cmdFirst.Visible = False
    cmdBusiness.Visible = False
    cmdCar.Visible = False
    cmdBus.Visible = False
    cmdCoach.Visible = False
End Sub

Private Sub cmdContinue_Click()
    RachelHaney3.Visible = False
    RachelHaney4.Visible = True
    RachelHaney4.cmdContinue.Visible = False
End Sub

Private Sub cmdFirst_Click()
    Travel = 1
    Total = Total + (300 * People)
    picResults.Print
    picResults.Print "You chose to fly first class to your destination."
    cmdContinue.Visible = True
    cmdBusiness.Visible = False
    cmdCoach.Visible = False
    cmdCar.Visible = False
    cmdBus.Visible = False
    cmdFirst.Visible = False
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

