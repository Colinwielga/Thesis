VERSION 5.00
Begin VB.Form frmCruisePorts 
   BackColor       =   &H0000FFFF&
   Caption         =   "Cruise Ports"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   6960
      Picture         =   "CruisePorts.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   2040
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   120
      Picture         =   "CruisePorts.frx":4755
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   5
      Top             =   2040
      Width           =   4095
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Caribbean Home Page"
      Height          =   1335
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   3495
   End
   Begin VB.CommandButton cmdTortolla 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tortola "
      Height          =   1335
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   3375
   End
   Begin VB.CommandButton cmdStThomas 
      BackColor       =   &H00C0FFC0&
      Caption         =   "St.Thomas"
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label lblghdjhs 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"CruisePorts.frx":A9AA
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblCruisePorts 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Cruise Ports"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmCruisePorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmCruisePorts
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form includes two command buttons that represent the different ports that the cruise ship
'will be stopping at throughout the trip and these buttons will bring you to one of two forms.

Private Sub cmdReturn_Click()
frmCruisePorts.Hide
frmCaribbeanHome.Show
End Sub

Private Sub cmdStThomas_Click()
frmStThomas.Show
frmCruisePorts.Hide
End Sub

Private Sub cmdTortolla_Click()
frmTortolla.Show
frmCruisePorts.Hide
End Sub

