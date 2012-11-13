VERSION 5.00
Begin VB.Form frmActivites 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   855
      Left            =   10680
      TabIndex        =   14
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Total"
      Height          =   2055
      Left            =   10680
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdcontinue 
      Caption         =   "Continue"
      Height          =   1935
      Left            =   10680
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   10680
      TabIndex        =   11
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Go to a Show $200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8160
      Picture         =   "frmActivites.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdspa 
      Caption         =   "Day at the Spa $250"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8160
      Picture         =   "frmActivites.frx":0B93
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdshopping 
      Caption         =   "Shopping Tour $175"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8160
      Picture         =   "frmActivites.frx":3A1E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdtour 
      Caption         =   "Tour Bus $85"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8160
      Picture         =   "frmActivites.frx":48AE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   7095
      Left            =   2760
      ScaleHeight     =   7035
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   1560
      Width           =   5175
   End
   Begin VB.CommandButton cmdskydiving 
      Caption         =   "SkyDiving $325"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      Picture         =   "frmActivites.frx":876A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdfishing 
      Caption         =   "Fishing $125"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      Picture         =   "frmActivites.frx":928F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdhiking 
      Caption         =   "Hiking $75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      Picture         =   "frmActivites.frx":DB94
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdhorse 
      Caption         =   "Horseback Riding $150"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      Picture         =   "frmActivites.frx":E5DF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblprices 
      Caption         =   " All Prices are Per Person."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label lbltilte 
      Caption         =   " Click the buttons to choose your activities."
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmActivites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmActivites
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/6/08
'Objective: Here we allow the user to select multiple activities while on vacation
'We have a running total that will display the total for all their activites when they are done.
'They do not have to select any activites before they move on

Option Explicit
Dim people As Integer
Dim subtotal As Single


Private Sub cmdclear_Click()

'Here the user has the option of clearing the activiites window and starting over

picResults.Cls

End Sub

Private Sub cmdcontinue_Click()

'Here the user goes to the next screen

frmActivites.Hide
frmairline.Show

End Sub

Private Sub cmdfishing_Click()

'Here the user selects an activity
'The user inputs how many people will participate in the activity
'The activity, number of people going, and cost are displayed
'The total for this activity is added to a running total

people = InputBox("Please enter the number of people participating in this activity.")

subtotal = people * 125

picResults.Print "Your party of "; people; " people is going fishing, for "; FormatCurrency(subtotal); "."

runningtotal = runningtotal + subtotal
End Sub

Private Sub cmdhiking_Click()

'Here the user selects an activity
'The user inputs how many people will participate in the activity
'The activity, number of people going, and cost are displayed
'The total for this activity is added to a running total

people = InputBox("Please enter the number of people participating in this activity.")

subtotal = people * 75

picResults.Print "Your party of "; people; " people is going hiking, for "; FormatCurrency(subtotal); "."

runningtotal = runningtotal + subtotal

End Sub

Private Sub cmdhorse_Click()

'Here the user selects an activity
'The user inputs how many people will participate in the activity
'The activity, number of people going, and cost are displayed
'The total for this activity is added to a running total

people = InputBox("Please enter the number of people participating in this activity.")

subtotal = people * 150

picResults.Print "Your party of "; people; " people is going horsebackriding, for "; FormatCurrency(subtotal); "."

runningtotal = runningtotal + subtotal

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdshopping_Click()

'Here the user selects an activity
'The user inputs how many people will participate in the activity
'The activity, number of people going, and cost are displayed
'The total for this activity is added to a running total

people = InputBox("Please enter the number of people participating in this activity.")

subtotal = people * 175

picResults.Print "Your party of "; people; " people is going on a shopping tour, for "; FormatCurrency(subtotal); "."

runningtotal = runningtotal + subtotal

End Sub

Private Sub cmdshow_Click()

'Here the user selects an activity
'The user inputs how many people will participate in the activity
'The activity, number of people going, and cost are displayed
'The total for this activity is added to a running total

people = InputBox("Please enter the number of people participating in this activity.")

subtotal = people * 200

picResults.Print "Your party of "; people; " people is going to a show, for "; FormatCurrency(subtotal); "."

runningtotal = runningtotal + subtotal

End Sub

Private Sub cmdskydiving_Click()

'Here the user selects an activity
'The user inputs how many people will participate in the activity
'The activity, number of people going, and cost are displayed
'The total for this activity is added to a running total

people = InputBox("Please enter the number of people participating in this activity.")

subtotal = people * 325

picResults.Print "Your party of "; people; " people is going skydiving, for "; FormatCurrency(subtotal); "."

runningtotal = runningtotal + subtotal

End Sub

Private Sub cmdspa_Click()

'Here the user selects an activity
'The user inputs how many people will participate in the activity
'The activity, number of people going, and cost are displayed
'The total for this activity is added to a running total

people = InputBox("Please enter the number of people participating in this activity.")

subtotal = people * 250

picResults.Print "Your party of "; people; " people is spending a day at the spa, for "; FormatCurrency(subtotal); "."

runningtotal = runningtotal + subtotal

End Sub

Private Sub cmdtotal_Click()

picResults.Print "***********************************************************************"
picResults.Print "Your total for your activities is "; FormatCurrency(runningtotal); "."
End Sub

Private Sub cmdtour_Click()

'Here the user selects an activity
'The user inputs how many people will participate in the activity
'The activity, number of people going, and cost are displayed
'The total for this activity is added to a running total

people = InputBox("Please enter the number of people participating in this activity.")

subtotal = people * 85

picResults.Print "Your party of "; people; " people is going sightseeing on a tour bus , for "; FormatCurrency(subtotal); "."

runningtotal = runningtotal + subtotal
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
