VERSION 5.00
Begin VB.Form CaseValue 
   Caption         =   "Form2"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8865
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   7815
      Left            =   4680
      Picture         =   "DoND2.frx":0000
      ScaleHeight     =   7755
      ScaleWidth      =   5355
      TabIndex        =   4
      Top             =   600
      Width           =   5415
   End
   Begin VB.PictureBox picResults5 
      BackColor       =   &H80000009&
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6915
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCaseValues 
      BackColor       =   &H80000002&
      Caption         =   "Display the Values That Are In the 26 Cases"
      Height          =   3855
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdBacktoStart 
      BackColor       =   &H80000013&
      Caption         =   "Back to Main Menu"
      Height          =   1095
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label lblMiniGame 
      BackColor       =   &H80000015&
      Caption         =   "Deal or No Deal Case Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "CaseValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Deal or No Deal Introduction
'Form Name: Start
'Authors: Chris Bergstrom and Brady King
'Date Written: March 27th, 2008
'Objective of Form: To allow the user to print a listing
'of the monetary values that will be randomly assigned
'to each of the 26 cases, so they know how much money
'can be won in the game.

Private Sub cmdBacktoStart_Click() 'This button takes the user back to the main menu
CaseValue.Hide
Start.Show
End Sub

Private Sub cmdCaseValues_Click() 'Allows the user to print the monetary values in the cases
picResults5.Print "These are the values randomly"
picResults5.Print "assigned to the 26 cases every "
picResults5.Print "show(Note: One value per case):"
picResults5.Print "$.01"
picResults5.Print "$1"
picResults5.Print "$5"
picResults5.Print "$10"
picResults5.Print "$25"
picResults5.Print "$50"
picResults5.Print "$75"
picResults5.Print "$100"
picResults5.Print "$200"
picResults5.Print "$300"
picResults5.Print "$400"
picResults5.Print "$500"
picResults5.Print "$750"
picResults5.Print "$1,000"
picResults5.Print "$5,000"
picResults5.Print "$10,000"
picResults5.Print "$25,000"
picResults5.Print "$50,000"
picResults5.Print "$75,000"
picResults5.Print "$100,000"
picResults5.Print "$200,000"
picResults5.Print "$300,000"
picResults5.Print "$400,000"
picResults5.Print "$500,000"
picResults5.Print "$750,000"
picResults5.Print "$1,000,000"

End Sub
