VERSION 5.00
Begin VB.Form frmVeniceAttractions 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Add Attractions to Budget"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   6480
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   255
   End
   Begin VB.PictureBox Picture5 
      Height          =   3015
      Left            =   6720
      Picture         =   "frmVeniceAttractions.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
   Begin VB.PictureBox Picture4 
      Height          =   2895
      Left            =   7080
      Picture         =   "frmVeniceAttractions.frx":2DFC6
      ScaleHeight     =   2835
      ScaleWidth      =   3075
      TabIndex        =   3
      Top             =   5400
      Width           =   3135
   End
   Begin VB.PictureBox Picture3 
      Height          =   2655
      Left            =   360
      Picture         =   "frmVeniceAttractions.frx":55A70
      ScaleHeight     =   2595
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   5280
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   4440
      Picture         =   "frmVeniceAttractions.frx":72B72
      ScaleHeight     =   3195
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   240
      Picture         =   "frmVeniceAttractions.frx":8D658
      ScaleHeight     =   2715
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Tour St. Mark's Basilica FREE"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   5160
      TabIndex        =   14
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Tour the Bell Tower $8"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   4800
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "2 Hour Walking Tour $27"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ride a Gondola for 20 minutes $50"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   7440
      TabIndex        =   8
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "See Grand Canal $5"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Attractions In Venice"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   855
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmVeniceAttractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmVeniceAttractions
'Jessica Florek
'Written: 3/12/09
'Objective: This form has a variety of options for entertainment in Venice that the user can choose
'and will be added to their budgets.


Option Explicit

'the check boxes determing what prices are subtracted from the budget and added to the veniceattractioncost that is later used and displayed in the budget summary

Private Sub cmdAdd_Click()
If Check4 Then
    budget = budget - 8
    veniceattractioncost = veniceattractioncost + 8
End If
If Check1 Then
    budget = budget - 5
    veniceattractioncost = veniceattractioncost + 5
End If
If Check2 Then
    budget = budget - 50
    veniceattractioncost = veniceattractioncost + 50
End If
If Check3 Then
    budget = budget - 27
    veniceattractioncost = veniceattractioncost + 27
End If

venice = True
'makes the information about venice displayed in the budget summary
frmVeniceAttractions.Hide
frmVenice.Show

End Sub

Private Sub cmdBack_Click()
frmVeniceAttractions.Hide
frmVenice.Show

End Sub

