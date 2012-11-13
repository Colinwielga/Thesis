VERSION 5.00
Begin VB.Form frmKetchikan 
   BackColor       =   &H0080C0FF&
   Caption         =   "Ketchikan"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFC0FF&
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
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Alaskan Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Compute total for all activities"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox txtHiking 
      Height          =   735
      Left            =   3360
      TabIndex        =   11
      Top             =   6120
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      Height          =   4215
      Left            =   4800
      ScaleHeight     =   4155
      ScaleWidth      =   5475
      TabIndex        =   8
      Top             =   3240
      Width           =   5535
   End
   Begin VB.TextBox txtMountainBuggying 
      Height          =   735
      Left            =   3360
      TabIndex        =   7
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtKayak 
      Height          =   735
      Left            =   3360
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtSportsFishing 
      Height          =   735
      Left            =   3360
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblfhjkdf 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hiking/Bike Rides: $50 per person"
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
      Left            =   240
      TabIndex        =   10
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label lbldfhksjd 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Ketchikan.frx":0000
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Label lblfhjsdfs 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mountain Buggying (Jeep Tours): $225 per person"
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
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label lblhjkshkd 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Kayak Expeditions: $105 per person"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label lbhfjksd 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sports Fishing: $300 per person"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label ldkjfio 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Ketchikan.frx":00BC
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label lblKetchikan 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ketchikan"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmKetchikan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmKetchikan
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form provides information regarding different types of activites the user could do if he/she
'decided to utilize the time at Ketchikan when the ship stops here and prices for those activities. There are
'text boxes for the user to enter their information as to how many, if any, people desire to participate in any
'of the activities listed. There is also a command button that computes the user's information and gives them a
'grand total of how much everything they have entered into the text boxes will end up costing them if they choose
'to engage in those activities.
Option Explicit

Private Sub cmdClear_Click()
picResults.Cls
End Sub

Private Sub cmdCompute_Click()
Dim runningtotal As Single, SportsFishing As Integer, Kayak As Integer, Hiking As Integer
Dim MountainBuggying As Integer, SportsFishingSum As Single, KayakSum As Single, MountainBuggyingSum As Single
Dim HikingSum As Single
runningtotal = 0

SportsFishing = txtSportsFishing.Text
Kayak = txtKayak.Text
MountainBuggying = txtMountainBuggying.Text
Hiking = txtHiking.Text


SportsFishingSum = SportsFishing * 300
picResults.Print "For"; SportsFishing; "person(s) to go sports fishing, the cost is "; FormatCurrency(SportsFishingSum, 2)
    runningtotal = SportsFishingSum + runningtotal

KayakSum = Kayak * 105
picResults.Print "For"; Kayak; "person(s) to kayak, the cost is "; FormatCurrency(KayakSum, 2)
    runningtotal = KayakSum + runningtotal

MountainBuggyingSum = MountainBuggying * 225
picResults.Print "For"; MountainBuggying; "person(s) to go mountain buggying, the cost is "; FormatCurrency(MountainBuggyingSum, 2)
    runningtotal = MountainBuggyingSum + runningtotal
    
HikingSum = Hiking * 50
picResults.Print "For"; Hiking; "person(s) to go hiking or biking, the cost is "; FormatCurrency(HikingSum, 2)
    runningtotal = HikingSum + runningtotal
picResults.Print "************************************************************************"
picResults.Print "The total for all your activities is "; FormatCurrency(runningtotal, 2)
    
End Sub

Private Sub cmdReturn_Click()
frmKetchikan.Hide
frmAlaskanHome.Show
End Sub

