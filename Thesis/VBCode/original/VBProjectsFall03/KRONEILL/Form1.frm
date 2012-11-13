VERSION 5.00
Begin VB.Form MainMenu 
   BackColor       =   &H00FF8080&
   Caption         =   "Welcome to Boeing Aircraft Company! "
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "Perpetua Titling MT"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8280
      TabIndex        =   4
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdAircraftPics 
      Caption         =   "Aircraft Images"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5880
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdAircraftPrice 
      Caption         =   "Aircraft Pricing"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3480
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdAircraftSpecs 
      Caption         =   "Aircraft Specifications"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1080
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
   End
   Begin VB.PictureBox pbxlogo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   1680
      Width           =   10455
   End
   Begin VB.Label lblauthordate 
      BackColor       =   &H00FF8080&
      Caption         =   "VB design by Kerry O'Neill  10/24/2003"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   7920
      Width           =   4335
   End
   Begin VB.Label lbldatabase 
      BackColor       =   &H00FF8080&
      Caption         =   "AirCRAFT INFORMATION DATABASE"
      Height          =   855
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   8415
   End
   Begin VB.Label lblboeing 
      BackColor       =   &H00FF8080&
      Caption         =   "BOEING COMPANY "
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project is designed to allow airline owners a promotional
'program that allows them to see specifications, pricing, and
'pictures of Boeing's modern airliner fleet. This will allow the
'buyers to make informed decisions on whether or not to purchase
'their aircraft from Boeing or another rival company/conglomerate
'such as Airbus Industrie.

'The purpose of this form is to serve as the start up form and provide
'navigation to any of the three other forms or to allow the user to exit
'the program.

Option Explicit
Private Sub cmdAircraftSpecs_Click() 'goes to AircraftSpecs(form2)
    MainMenu.Hide
    AircraftSpecs.Show
    MsgBox ("On this page you will find important specifications and facts about our line of aircraft. Please use the 'read and print' button first to load the information from an outside file.") 'this will pop up before form two is displayed
End Sub
Private Sub cmdAircraftPrice_Click() 'goes to AircraftPrice(form3)
    MainMenu.Hide
    AircraftPrice.Show
    MsgBox ("This page will provide you with minimum pricing for our aircraft. The base price model does not include any upgrades you may wish to have installed prior to delivery. Please use the 'read and print' button first to load the information from an outside file.") 'this will pop up before form 3 is displayed
End Sub
Private Sub cmdAircraftPics_Click() 'goes to AircraftPics(form4)
    MainMenu.Hide
    AircraftPics.Show
    MsgBox ("This page will allow you to view an image of each of our production aircraft to help guide you in your decision-making process.") 'this will pop up before form 4 is displayed
End Sub

Private Sub cmdquit_Click() 'ends program
    MsgBox ("Thank you for considering a fine Boeing aircraft as a possible investment for your company. Please contact a Boeing representative to begin the buying process.") 'this will pop up before the program ends
    End
End Sub

