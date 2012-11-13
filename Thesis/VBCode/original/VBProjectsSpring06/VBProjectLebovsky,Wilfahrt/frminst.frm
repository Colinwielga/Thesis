VERSION 5.00
Begin VB.Form frminst 
   BackColor       =   &H00000000&
   Caption         =   "Instructions"
   ClientHeight    =   6705
   ClientLeft      =   2760
   ClientTop       =   2280
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9765
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   1335
      Left            =   6960
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to main"
      Height          =   1335
      Left            =   480
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdnav 
      BackColor       =   &H00000000&
      Caption         =   "Command2"
      Height          =   2895
      Left            =   4800
      Picture         =   "frminst.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton cmdrace 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How to Race"
      Height          =   3135
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frminst.frx":3072
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Clay Wilfahrt and Andy Lebovsky"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   6120
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Click for help navigating and using this program"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Click to learn how to Race"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frminst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frminst(frminst.frm)
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'The Objective of this form is to explain how to use the racing format and also to help navigate throughout the project

Option Explicit
'Exits the Program
Private Sub cmdexit_Click()
End
End Sub
'Gives info on how to navigate throughout the program
Private Sub cmdnav_Click()
    MsgBox "At the main menu, there are three options.  The 'about' button displays information about the program as well as the creators of the project.  The 'instructions' button will lead you to a page offering help on how to begin the race as well as how to navigate the site.  The 'race' button will take you to a screen allowing you to select a car.  At the select car screen, you have the option of either selecting a car and begginning the race, or returning to the main menu.  At any time you can select the 'exit' button to exit the program."
End Sub
'Gives instructions on how to race
Private Sub cmdrace_Click()
    MsgBox "At the main menu, select the 'race' button.  This will bring you to a screen where you will select a car of your choice.  You can either select a car, or click the 'stats' button to review the specs of each car.  Once a car has been selected, continue to the actual race via the 'Begin Race!' button.  To begin the race, click the 'start race' button.  A question will come up, and if answered correctly, your selected car will move forward towards the finish line.  If you answer incorrectly, your car will not move.  First car to fully cross the finish line wins.  Be sure to review all of the information on the site before racing, if you don't, the computer might win!", , "Racing Instructions"
End Sub
'Brings you back to main screen
Private Sub cmdreturn_Click()
    frminst.Hide
    frmmain.Show
End Sub


