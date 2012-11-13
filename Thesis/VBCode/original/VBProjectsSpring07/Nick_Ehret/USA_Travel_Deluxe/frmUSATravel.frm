VERSION 5.00
Begin VB.Form frmUSATravel 
   Caption         =   "USA Travel Planner"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   975
      Left            =   1560
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   4320
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Find Total Cost"
      Height          =   975
      Left            =   5640
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdExpenses 
      Caption         =   "Other Trip Expenses"
      Height          =   975
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdTravel 
      Caption         =   "Travel Expenses"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblHome 
      Alignment       =   2  'Center
      Caption         =   "USA Travel Planner Deluxe"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "frmUSATravel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub cmdBack_Click()
    'This button will go back to frmCity
    
    frmUSATravel.Visible = False
    frmCity.Visible = True
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdExpenses_Click()
    'This button will bring the user to frmOtherExp
    
    frmUSATravel.Visible = False
    frmOtherExp.Visible = True
    
End Sub

Private Sub cmdTotal_Click()
    'This button will bring the user to frmTotalCost
    
    frmUSATravel.Visible = False
    frmTotalCost.Visible = True
    

End Sub

Private Sub cmdTravel_Click()
    'This button will bring the user to frmTravel
    
    frmUSATravel.Visible = False
    frmTravel.Visible = True
    
End Sub
