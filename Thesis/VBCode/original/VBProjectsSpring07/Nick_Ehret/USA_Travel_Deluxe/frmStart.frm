VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H000000C0&
   Caption         =   "Start Your Trip"
   ClientHeight    =   5535
   ClientLeft      =   1500
   ClientTop       =   1290
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7335
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   4080
      TabIndex        =   3
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "OK"
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtTraveler 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "USA Travel Planner Deluxe"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "To start planning your weeklong USA getaway from Minnesota please enter your name"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   5775
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: USA Travel Planner Deluxe
'Written By: Nick Ehret
'Date Written: March 21, 2007
'Project Purpose: The overall purpose of my project is too allow a person traveling out'
                'of minneapolis to plan a weeklong trip in a variety of ways. My program'
                'allows the user to choose multiple options such as whether to fly or drive,'
                'how much to spend on food and a hotel, and whether they want to rent a car.'
                'My program also then compares the actual cost of the trip to the traveler's
                'preferred cost and then shows the actual cost of a trip in other cities.
                'Overall, my program is meant to help plan a trip around the USA.
'Form Purpose: The purpose of this form is for the user to enter their name into a text box'
                'and then that name is stored throughout the program and used elsewhere.
                'This form also brings the user to frmCity
                
Option Explicit

Private Sub cmdExit_Click()
 End 'this will end the program
End Sub

Private Sub cmdName_Click()
    'this button will move to frmCity so a city number can be entered and will store,
    'Name in the module
    
    Traveler = txtTraveler.Text
    
    frmStart.Visible = False 'makes frmStart not visible
    frmCity.Visible = True  'makes frmCity visible
    
    
    
End Sub
