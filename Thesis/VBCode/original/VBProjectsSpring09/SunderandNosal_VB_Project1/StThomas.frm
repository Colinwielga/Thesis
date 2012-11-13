VERSION 5.00
Begin VB.Form frmStThomas 
   BackColor       =   &H80000003&
   Caption         =   "St. Thomas"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Caribbean Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H0080FF80&
      Caption         =   "Compute total cost"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   3975
      Left            =   5760
      ScaleHeight     =   3915
      ScaleWidth      =   4515
      TabIndex        =   6
      Top             =   3000
      Width           =   4575
   End
   Begin VB.TextBox txtParasailing 
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtSwimmingwithDolphins 
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtSnorkeling 
      Height          =   735
      Left            =   3960
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblParasailing 
      BackColor       =   &H00FFFF00&
      Caption         =   "Parasailing: $100 per person"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label lblSwimmingwithDolphins 
      BackColor       =   &H00FFFF00&
      Caption         =   "Swimming with Dolphins: $75 per person"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblSnorkeling 
      BackColor       =   &H00FFFF00&
      Caption         =   "Snorkeling: $25.95 per person"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"StThomas.frx":0000
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   7455
   End
   Begin VB.Label lblStThomas 
      BackColor       =   &H00FFFFC0&
      Caption         =   " St. Thomas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmStThomas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmStThomas
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form provides information regarding different types of activites the user could do if he/she
'decided to utilize the time at St. Thomas when the ship stops here and prices for those activities. There are
'text boxes for the user to enter their information as to how many, if any, people desire to participate in any
'of the activities listed. There is also a command button that computes the user's information and gives them a
'grand total of how much everything they have entered into the text boxes will end up costing them if they choose
'to engage in those activities.

Option Explicit

Private Sub cmdClear_Click()
picResults.Cls

End Sub

Private Sub cmdCompute_Click()
Dim runningtotal As Single, Parasailing As Integer, Snorkeling As Integer
Dim Swimming As Integer, SnorkelingSum As Single, SwimmingSum As Single, ParasailingSum As Single

runningtotal = 0

Parasailing = txtParasailing.Text
Snorkeling = txtSnorkeling.Text
Swimming = txtSwimmingwithDolphins.Text

SnorkelingSum = Snorkeling * 25.95
picResults.Print "For"; Snorkeling; "people to snorkel, the cost is "; FormatCurrency(SnorkelingSum, 2)
    runningtotal = SnorkelingSum + runningtotal

SwimmingSum = Swimming * 75
picResults.Print "For"; Swimming; "people to swim with the dolphins, the cost is "; FormatCurrency(SwimmingSum, 2)
    runningtotal = SwimmingSum + runningtotal

ParasailingSum = Parasailing * 100
picResults.Print "For"; Parasailing; "people to parasail, the cost is "; FormatCurrency(ParasailingSum, 2)
    runningtotal = ParasailingSum + runningtotal
    
picResults.Print "******************************************************"
    
picResults.Print "The total for all your activities is "; FormatCurrency(runningtotal, 2)
    
End Sub

Private Sub cmdReturn_Click()
frmStThomas.Hide
frmCaribbeanHome.Show
End Sub
