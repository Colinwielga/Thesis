VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Stat Tracker"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmTitle.frx":0000
   ScaleHeight     =   6705
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Back to Main Menu"
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   5880
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   5880
      TabIndex        =   2
      Top             =   5880
      Width           =   3375
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H000080FF&
      Caption         =   "Create New Data"
      Height          =   735
      Left            =   5880
      MaskColor       =   &H000080FF&
      TabIndex        =   1
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H000080FF&
      Caption         =   "Load Existing Data"
      Height          =   735
      Left            =   480
      MaskColor       =   &H000080FF&
      TabIndex        =   0
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Stat Tracker"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   54.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: GolfGuide
':Form Name:  frmStats
':Author:   Tyler Cash
':Date written:  March 21, 2009


'This is the title screen for the Stat Tracking portion of the program.
'This form allows the user to indicate whether they will be creating a new text file
'or loading an pre-existing text file.
'The user is also able to go back the the main menu.

Option Explicit

Private Sub cmdCreate_Click()
'This button changes forms to the form allowing the user to create a new text file.
    
'Changing forms
    frmStats.Hide
    frmCreate.Show
End Sub

Private Sub cmdExit_Click()
'This button changes forms to the main menu form.
    
'Changing forms
    frmStats.Hide
    frmTitle.Show
End Sub

Private Sub cmdLoad_Click()
'This button changes forms to the form alling the user to load a text file.

'Changing forms
    frmLoad.Show
    frmStats.Hide
End Sub

Private Sub cmdQuit_Click()
'This button ends the program

    End
End Sub

