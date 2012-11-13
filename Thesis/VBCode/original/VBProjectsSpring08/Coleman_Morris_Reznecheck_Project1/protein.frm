VERSION 5.00
Begin VB.Form frmprotein 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   FillColor       =   &H80000012&
   LinkTopic       =   "Form2"
   Picture         =   "protein.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Go Back to Main Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      MaskColor       =   &H0080FFFF&
      TabIndex        =   6
      Top             =   720
      Width           =   3135
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   240
      ScaleHeight     =   5235
      ScaleWidth      =   6075
      TabIndex        =   5
      Top             =   3360
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Common Foods and the Amount of Protein Contained"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      MaskColor       =   &H0080FFFF&
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton cmdprotein 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compute Healthy Protein Intake"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   13200
      MaskColor       =   &H0080FFFF&
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   8160
      ScaleHeight     =   4035
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   3600
      Width           =   7095
   End
   Begin VB.TextBox txtweight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      TabIndex        =   0
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "INPUT WEIGHT   (IN POUNDS)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
End
Attribute VB_Name = "frmprotein"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'Form frmprotein
'Joel Coleman
'March 29, 2008
'To assist people in figuring out a healthy amount of protein to consume in relation with goal such as dieting or bodybuilding
Option Explicit
Dim weight As Integer, lowdiet As Integer, highdiet As Integer, lowps As Single, highps As Single
Dim lowsb As Single, highsb As Single, lowe As Single, highe As Single
Private Sub cmdprotein_Click()
'stores input weight in order to do calculations with it
weight = txtWeight.Text

'Got the equation from this site: http://www.sharpweblabs.com/shop/heart.htm
picResults.Print "Your Current Protein Needs:"
picResults.Print "Goal", Tab(25), "Range(Grams)"
'calculates equation and prints
lowdiet = 0.35 * weight
highdiet = weight * 1
picResults.Print "On a Diet:", Tab(30); lowdiet, Tab(40), highdiet
'calculates equation and prints
lowps = weight * 0.9
highps = weight * 1.1
picResults.Print "Power & Speed", Tab(30); lowps, Tab(40); highps
'calculates equation and prints
lowsb = weight * 1
highsb = weight * 1.6
picResults.Print "Strength & Bodybuilding", Tab(30); lowsb, Tab(40); highsb
'calculates equation and prints
lowe = weight * 0.7
highe = weight * 0.9
picResults.Print "Endurance", Tab(30); lowe, Tab(40); highe



End Sub

Private Sub Command1_Click()

Dim food(1 To 20) As String, grams(1 To 20) As String, CTR As Integer
CTR = 0
'Got information from http://www.annecollins.com/protein-foods.htm
Open App.Path & "\proteinfoods.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, food(CTR), grams(CTR)
    Loop
Close #1
    
For CTR = 1 To 15
picOutput.Print food(CTR); Tab(25); grams(CTR)
Next CTR
End Sub

Private Sub Command2_Click()
frmMainpage.Show
frmprotein.Hide
End Sub

