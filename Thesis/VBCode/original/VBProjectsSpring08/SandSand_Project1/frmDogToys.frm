VERSION 5.00
Begin VB.Form frmDogToys 
   BackColor       =   &H00000080&
   Caption         =   "Dog Toys"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H0080FFFF&
      Caption         =   "Move to Dog Food"
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   4095
   End
   Begin VB.CommandButton cmdNo 
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
      Left            =   8760
      TabIndex        =   5
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11040
      TabIndex        =   3
      Top             =   9360
      Width           =   3975
   End
   Begin VB.CommandButton cmdBone 
      BackColor       =   &H0080FFFF&
      Caption         =   "Bone"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton cmdTennisBall 
      BackColor       =   &H0080FFFF&
      Caption         =   "Tennis Ball"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   3615
   End
   Begin VB.CommandButton cmdRope 
      BackColor       =   &H0080FFFF&
      Caption         =   "Rope"
      BeginProperty Font 
         Name            =   "Minion Pro Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label labInstructions 
      BackColor       =   &H00000080&
      Caption         =   "Click on the buttons to select what toys you would like for your dog (you can choose more than one).  "
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
      Height          =   1215
      Left            =   960
      TabIndex        =   7
      Top             =   720
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Dog Toys"
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
      Height          =   1935
      Left            =   1560
      TabIndex        =   4
      Top             =   7680
      Width           =   4335
   End
End
Attribute VB_Name = "frmDogToys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmDogToys
'Author: Scott Sand and Kate Sand
'Date Written: March 12, 2008
'Objective: This is where people can select toys for their dogs.
'Other Comments:
Option Explicit

Private Sub cmdBack_Click()
'The customer is directed back to the main menu
frmDogToys.Hide
frmMainMenu.Show
End Sub

Private Sub cmdBone_Click()
'The customer purchases a dog bone
MsgBox ("You have chosen to purchase a Bone that costs $4")
AccesoriesCost = AccesoriesCost + 4
End Sub


Private Sub cmdMove_Click()
'Moves the customer to the dog food when they are finished purchasing toys
frmDogToys.Hide
frmDogFood.Show
End Sub

Private Sub cmdNo_Click()
'The customer can choose not to purchase dog toys
frmDogToys.Hide
frmDogFood.Show
End Sub

Private Sub cmdRope_Click()
'The customer purchases the rope
MsgBox ("You have chosen to purchase a Play Rope that costs $3")
AccesoriesCost = AccesoriesCost + 3
End Sub

Private Sub cmdTennisBall_Click()
'The customer purchases the tennis ball
MsgBox ("You have chosen to purchase a Tennis Ball that costs $2")
AccesoriesCost = AccesoriesCost + 2
End Sub
