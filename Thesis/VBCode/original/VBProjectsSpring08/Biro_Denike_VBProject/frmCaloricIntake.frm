VERSION 5.00
Begin VB.Form frmCaloricIntake 
   BackColor       =   &H000080FF&
   Caption         =   "Calculate your caloric intake"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Main Screen"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   3075
      TabIndex        =   7
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtActivity 
      Height          =   615
      Left            =   5640
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblActivity 
      BackColor       =   &H000080FF&
      Caption         =   "Enter your activity level here ==> (light, moderate, or heavy)"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H000080FF&
      Caption         =   "Activity Level Descriptions"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label lblHeavy 
      BackColor       =   &H000080FF&
      Caption         =   "Heavy= athletes"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblModerate 
      BackColor       =   &H000080FF&
      Caption         =   "Moderate= active in addition to workout"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblLight 
      BackColor       =   &H000080FF&
      Caption         =   "Light= 30 minute workout/day"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmCaloricIntake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Bon Appetit:Menu Planner
'Form name: Caloric Intake (frmCaloricIntake.frm)
'Authors: Sarah Biro and Heather Denike
'Date written: 3/13/2008
'Objective: This form gets weight and activity level as input from the user.
            'It then calculates the recommended range of calories to consume in a day.

Option Explicit




Private Sub cmdcalculate_Click()
    Dim pounds As Single, activity As String, weight As Integer
    Dim bmr As Single

    pounds = InputBox("Enter your weight in pounds.") 'user inputs weight
    weight = pounds / 2.2
    bmr = weight * 24
    activity = txtActivity.Text 'user inputs activity type in textbox
    
    'calculates caloric intake range
    If activity = "light" Then
        low = 0.5 * bmr + bmr
        High = 0.7 * bmr + bmr
    ElseIf activity = "moderate" Then
        low = 0.65 * bmr + bmr
        High = 0.8 * bmr + bmr
    Else: activity = "heavy"
        low = 0.9 * bmr + bmr
        High = 1.2 * bmr + bmr
    End If
    
    'prints caloric intake range
    picResults.Print " Your recommended daily caloric "
    picResults.Print " intake ranges from "; FormatNumber(low, 0); " to "; FormatNumber(High, 0); "."

End Sub

Private Sub cmdReturn_Click()
    'returns to main form
    frmmain.Show
    frmCaloricIntake.Hide
End Sub

