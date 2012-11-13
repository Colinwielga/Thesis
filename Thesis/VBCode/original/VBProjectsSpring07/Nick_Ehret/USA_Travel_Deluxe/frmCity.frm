VERSION 5.00
Begin VB.Form frmCity 
   BackColor       =   &H00C00000&
   Caption         =   "Pick Your Travel City"
   ClientHeight    =   5940
   ClientLeft      =   885
   ClientTop       =   1080
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2697.548
   ScaleMode       =   0  'User
   ScaleWidth      =   8385
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   855
      Left            =   3000
      TabIndex        =   18
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdTrip 
      Caption         =   "Compute the Costs of the Trip"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   5760
      TabIndex        =   0
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblWashDC 
      Caption         =   "12-Washington D.C."
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please click on a city from the list below that you wish to travel to from Minneapolis"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   285
      TabIndex        =   16
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label lblSanFran 
      BackColor       =   &H000000C0&
      Caption         =   "1-San Francisco"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblNewYork 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2-New York City"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblLosAngeles 
      BackColor       =   &H000000C0&
      Caption         =   "3-Los Angeles"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblChicago 
      Caption         =   "4-Chicago"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblOrlando 
      BackColor       =   &H000000C0&
      Caption         =   "5-Orlando"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblMiami 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6-Miami"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2843
      TabIndex        =   10
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000C0&
      Caption         =   "7-Seattle"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2843
      TabIndex        =   9
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblNewOrleans 
      Caption         =   "8-New Orleans"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2843
      TabIndex        =   8
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblDenver 
      BackColor       =   &H000000C0&
      Caption         =   "9-Denver"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2843
      TabIndex        =   7
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label10 
      Caption         =   "10-Nashville"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2843
      TabIndex        =   6
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblBoston 
      BackColor       =   &H000000C0&
      Caption         =   "11-Boston"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblAtlanta 
      BackColor       =   &H000000C0&
      Caption         =   "13-Atlanta"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblPhoenix 
      Caption         =   "14-Phoenix"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblLasVegas 
      BackColor       =   &H000000C0&
      Caption         =   "15-Las Vegas"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   4200
      Width           =   2415
   End
End
Attribute VB_Name = "frmCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: USA Travel Planner Deluxe
'Written By: Nick Ehret
'Date Written: March 21, 2007
'Form Purpose: 'The purpose of this form is to allow the user to pick a city that he or she'
                'wants to travel too by clicking on the label of that city. Once a city is'
                'picked then a msgbox shows the user that he or she has selected the city.
                'Then the user can click a button to move on to calculate the costs of the trip
                'but if the user hasn't selected a city they can't move on
                

Option Explicit

Private Sub cmdBack_Click()
    'This button will bring the user back to frmStart
    
    frmStart.Visible = True 'makes frmStart visible
    frmCity.Visible = False 'makes frmcity not visible
    
End Sub
Private Sub cmdExit_Click()
    End 'This button will end the program
End Sub

Private Sub cmdTrip_Click()
    'this button will move to frmUSATravel and put back and error message if,
    'a city has not been selected
    
    If CityNum > 0 And CityNum < 16 Then 'use an if statement to make sure the user has selected a city
        frmCity.Visible = False 'if the user has selected a city then go to frmtravel
        frmTravel.Visible = True
    Else
        MsgBox "Please choose a city you want to travel too", , "Error"
    End If
End Sub

Private Sub Label10_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Nashville"
    CityNum = 10
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub Label7_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Seattle"
    CityNum = 7
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblAtlanta_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Atlanta"
    CityNum = 13
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblBoston_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Boston"
    CityNum = 11
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblChicago_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Chicago"
    CityNum = 4
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblDenver_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Denver"
    CityNum = 9
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblLasVegas_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Las Vegas"
    CityNum = 15
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblLosAngeles_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Los Angeles"
    CityNum = 3
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblMiami_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Miami"
    CityNum = 6
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblNewOrleans_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "New Orleans"
    CityNum = 8
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblNewYork_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "New York City"
    CityNum = 2
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblOrlando_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Orlando"
    CityNum = 5
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblPhoenix_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Phoenix"
    CityNum = 14
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub

Private Sub lblSanFran_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "San Francisco"
    CityNum = 1
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
    
    
End Sub

Private Sub lblWashDC_Click()
    'If the user clicks on this label that will become the travel city and a messagebox,
    'will display that fact
    
    TravelCity = "Washington D.C."
    CityNum = 12
    MsgBox "You have chosen " & TravelCity & " as your travel city!"
End Sub
