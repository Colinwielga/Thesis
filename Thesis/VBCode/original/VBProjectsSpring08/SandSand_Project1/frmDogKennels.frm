VERSION 5.00
Begin VB.Form frmDogKennels 
   BackColor       =   &H00000080&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Dog Kennels"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOutsideCage 
      Caption         =   "Outside Cage"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   9
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlasticKennel 
      Caption         =   "Plastic Kennel"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   8
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdDogHouse 
      Caption         =   "Dog House"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdNoThanks 
      Caption         =   "No Thank You!"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      TabIndex        =   5
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10920
      TabIndex        =   4
      Top             =   9360
      Width           =   3975
   End
   Begin VB.PictureBox PicResults 
      Height          =   5295
      Left            =   4080
      ScaleHeight     =   5235
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   1680
      Width           =   6855
   End
   Begin VB.CommandButton cmdPlastic 
      BackColor       =   &H0080FFFF&
      Caption         =   "Plastic Kennel"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton cmdHouse 
      BackColor       =   &H0080FFFF&
      Caption         =   "Dog House"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton cmdCage 
      BackColor       =   &H0080FFFF&
      Caption         =   "Outside Cage"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label labInstructions 
      BackColor       =   &H00000080&
      Caption         =   "To view the types of dog kennels click below:"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   975
      Left            =   840
      TabIndex        =   11
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label LabInstructions3 
      BackColor       =   &H00000080&
      Caption         =   "Click the type of dog kennel you wish to purchase:"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Index           =   1
      Left            =   11280
      TabIndex        =   10
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Dog Kennels"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   48
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   960
      TabIndex        =   6
      Top             =   8280
      Width           =   5895
   End
End
Attribute VB_Name = "frmDogKennels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmDogKennels
'Author: Scott Sand and Kate Sand
'Date Written: March 10, 2008
'Objective: This is where customers can view and select a kennel for their dogs.
'Other Comments:

Option Explicit

Private Sub cmdCage_Click()
' A picture of the outside cage is displayed
Open App.Path & "\PicKennels.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Kennels(CTR)
Loop
PicResults.Picture = LoadPicture(Kennels(1))
Close #1
End Sub


Private Sub cmdDogHouse_Click()
'The customer can purchase the dog house
'A message box tells them how much it costs
MsgBox ("You have purchased a dog house for $76.00.")
HabitatCost = HabitatCost + 76
frmDogToys.Show
frmDogKennels.Hide
End Sub

Private Sub cmdHouse_Click()
'Customers can view a picture of the dog house
Open App.Path & "\PicKennels.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Kennels(CTR)
Loop
PicResults.Picture = LoadPicture(Kennels(3))
Close #1
End Sub

Private Sub cmdMainMenu_Click()
'Directs customers back to the main menu
frmMainMenu.Show
frmDogKennels.Hide
End Sub

Private Sub cmdNoThanks_Click()
'The customr can choose not to buy a kennelo for their dog
frmDogKennels.Hide
frmDogToys.Show
End Sub

Private Sub cmdOutsideCage_Click()
'The customer chooses to purchase the outside cage
MsgBox ("You have purchased a outside dog kennel for $149.00.")
HabitatCost = HabitatCost + 149
frmDogToys.Show
frmDogKennels.Hide
End Sub

Private Sub cmdPlastic_Click()
'The customer can view a picture of the plastic kennel
Open App.Path & "\PicKennels.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Kennels(CTR)
Loop
PicResults.Picture = LoadPicture(Kennels(2))
Close #1
End Sub


Private Sub cmdPlasticKennel_Click()
'The customer purchases the plastic kennel
MsgBox ("You have purchased a plastic kennel for $59.00.")
HabitatCost = HabitatCost + 59
frmDogToys.Show
frmDogKennels.Hide
End Sub
