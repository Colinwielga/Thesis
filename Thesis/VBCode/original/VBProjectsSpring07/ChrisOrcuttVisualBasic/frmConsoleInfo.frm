VERSION 5.00
Begin VB.Form frmConsoleInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Videogame Console Information"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenesis 
      Caption         =   "Sega Genesis"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdDreamcast 
      Caption         =   "Sega Dreamcast"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdPS1 
      Caption         =   "Sony Playstation"
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdPS2 
      Caption         =   "Sony Playstation 2"
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdPS3Prem 
      Caption         =   "Sony Playstation 3"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdPremium 
      Caption         =   "Microsoft Xbox 360"
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdXbox 
      Caption         =   "Microsoft Xbox"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdWii 
      Caption         =   "Nintendo Wii"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdGamecube 
      Caption         =   "Nintendo GameCube"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdN64 
      Caption         =   "Nintendo 64"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdSNES 
      Caption         =   "Super Nintendo"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdNES 
      Caption         =   "Nintendo Entertainment System"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblClick 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Console Pictures to Display Information"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   2640
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   3000
      Picture         =   "frmConsoleInfo.frx":0000
      Top             =   3240
      Width           =   3000
   End
End
Attribute VB_Name = "frmConsoleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmConsoleInfo
'26 March 2007
'This form lists 12 different videogame systems, each having their
'own command link to and a brief overview of each system.
'By clicking on the console pictures, users can display the
'summary/ overview descriptions.
Option Explicit
Private Sub cmdDreamcast_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmDreamcast.Show       'Shows Dreamcast form
End Sub
Private Sub cmdGamecube_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmGameCube.Show        'Shows GameCube form
End Sub
Private Sub cmdGenesis_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmSegaGenesis.Show     'Shows SegaGenesis form
End Sub
Private Sub cmdN64_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmN64.Show             'Shows N64 form
End Sub
Private Sub cmdNES_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmNES.Show             'Shows NES form
End Sub
Private Sub cmdPremium_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmXbox360.Show         'Shows Xbox360 form
End Sub
Private Sub cmdPS1_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmPS1.Show             'Shows PS1 form
End Sub
Private Sub cmdPS2_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmPS2.Show             'Shows PS2 form
End Sub
Private Sub cmdPS3Prem_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmPS3.Show             'Shows PS3 form
End Sub
Private Sub cmdReturn_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmSelectWant.Show      'Shows SelectWant form
End Sub
Private Sub cmdSNES_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmSuperNES.Show        'Shows SuperNES form
End Sub
Private Sub cmdWii_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmWii.Show             'Shows Wii form
End Sub
Private Sub cmdXbox_Click()
    frmConsoleInfo.Hide     'Hides ConsoleInfo form
    frmXbox.Show            'Shows Xbox form
End Sub
