VERSION 5.00
Begin VB.Form frmHome 
   Caption         =   "Form1"
   ClientHeight    =   11505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   11505
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H0000C000&
      Caption         =   "Program Information"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10800
      Width           =   2775
   End
   Begin VB.CommandButton cmdHunt 
      BackColor       =   &H000040C0&
      Caption         =   "DEER HUNTING"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton cmdDVC 
      BackColor       =   &H000040C0&
      Caption         =   "Deer and Vehicles"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton cmdTerms 
      BackColor       =   &H000040C0&
      Caption         =   "View deer specific terms."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblWhitetail 
      BackColor       =   &H0080FFFF&
      Caption         =   "Minnesota's Whitetail Deer"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MN Deer'
'Form Name: Home'
'Authors: Jordon Przybilla'
'Date Written: October 4, 2009
'this form will be the base for the rest of the program. the other forms will be accessed from here'
'also all the buttons on this form will load any arrays needed in the form that the button accesses'
Option Explicit

Private Sub cmdDVC_Click()
'takes user to Vehicles form and loads all arrays for the Vehicles form

Open App.Path & "\Data\CantAvoid.txt" For Input As #3
    Ctr = 0
        Do While Not EOF(3)
            Ctr = Ctr + 1
            Input #3, CantAvoid(Ctr)
        Loop
Close #3


Open App.Path & "\Data\DriveFacts.txt" For Input As #1
    Ctr = 0
        Do While Not EOF(1)
            Ctr = Ctr + 1
            Input #1, Facts(Ctr)
                
        Loop
Close #1

Open App.Path & "\Data\Avoid.txt" For Input As #2
    Ctr = 0
        Do While Not EOF(2)
            Ctr = Ctr + 1
            Input #2, AvoidTips(Ctr)
        Loop
Close #2

frmHome.Hide
frmVehicles.Show

End Sub

Private Sub cmdHunt_Click()
'this button take the user to the hunting form
frmHome.Hide
frmHunting.Show

End Sub

Private Sub cmdInfo_Click()
'this button will take the user to information about this program and load the information from a data file

Open App.Path & "\Data\Info.txt" For Input As #1
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, info(Ctr)
    Loop
Close #1

frmHome.Hide
frmInfo.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub


Private Sub cmdTerms_Click()
'this button will take the user to a form that will allow them to view deer termonology'
'also it will read the arrays for the Terms form so no extra buttons need to be pushed for the next form to work'

Open App.Path & "\Data\BuckDoeFawn.txt" For Input As #1
    Ctr = 0
        Do While Not EOF(1)
            Ctr = Ctr + 1
            Input #1, Terms(Ctr)
        Loop
Close #1

frmHome.Hide
frmTerms.Show

End Sub

