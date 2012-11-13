VERSION 5.00
Begin VB.Form frmLondonAttractions 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF00&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFF00&
      Caption         =   "Add Attractions to your Budget"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CheckBox chk4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   2175
      Left            =   2160
      Picture         =   "frmLondonAttractions.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2595
      TabIndex        =   9
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CheckBox chk3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   4200
      Picture         =   "frmLondonAttractions.frx":16932
      ScaleHeight     =   2475
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   3000
      Picture         =   "frmLondonAttractions.frx":2D4D4
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin VB.CheckBox chk2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3240
      TabIndex        =   0
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblPrivateEye 
      BackColor       =   &H00FF0000&
      Caption         =   "Dinner and Wine in Private Cart on the London Eye $350"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5760
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblGuards 
      BackColor       =   &H00FF0000&
      Caption         =   "Changing of the Guards FREE"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label lblBus 
      BackColor       =   &H00FF0000&
      Caption         =   "Day Tour on Double Decker Bus $16"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label lblEye 
      BackColor       =   &H00FF0000&
      Caption         =   "The London Eye  $12"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF0000&
      Caption         =   "London Attractions: Select the boxes next to the attractions you wish to see!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmLondonAttractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmLondonAttractions
'Jessica Florek
'Written: 3/6/09
'Objective: This form has a variety of options for entertainment in London that the user can choose
'and will be added to their budgets.


Option Explicit



Private Sub cmdAdd_Click()
'Depending on which checks are clicked on, the different options will be calculated into the budget and added to the londonattracioncost which will be used to help summarize the budget
If chk2 Then
    budget = budget - 12
    londonattractioncost = londonattractioncost + 12
End If
If chk1 Then
    budget = budget - 16
    londonattractioncost = londonattractioncost + 16
End If
If chk4 Then
    budget = budget - 350
    londonattractioncost = londonattractioncost + 350
End If
london = True
'this will cause the london summaries to be displayed in the budget summary

frmLondonAttractions.Hide
frmLondon.Show

End Sub

Private Sub cmdBack_Click()

frmLondonAttractions.Hide
frmLondon.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

