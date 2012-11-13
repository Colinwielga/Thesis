VERSION 5.00
Begin VB.Form frmFront 
   BackColor       =   &H00FF0000&
   Caption         =   "Dave and JB's Skate and Ski Store"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   Picture         =   "project.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNavitgate 
      Caption         =   "Navigate our Site"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdSki 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Ski Store"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   1080
      Picture         =   "project.frx":CC77
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdSkiWiziard 
      Caption         =   "Enter Skate Store"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   7680
      Picture         =   "project.frx":EA9B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label lblLayOut 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "We have two distinct Stores for you to choose from.  Click the appropriate link below to move to that store!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   840
      TabIndex        =   4
      Top             =   5640
      Width           =   5535
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Dave and JB's Skate and Ski Store Where We Get You What You Need!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   6720
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmFront"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dave and JB's ski and skate store
'frmFront
'Dave Lambert and John Brine
'Thursday March 23
'This form is used to have the user decide where they want to go



Private Sub cmdNavitgate_Click()
    'navigates to the navigate page
    frmNavigate.Visible = True
    frmFront.Visible = False
End Sub

Private Sub cmdSki_Click()
    'navigates to the ski store page
    frmFront.Hide
    frmSkiStore.Show
End Sub

Private Sub cmdSkiWiziard_Click()
    'navigates to the skate store page
    frmFront.Hide
    frmSkateStore.Show
End Sub

