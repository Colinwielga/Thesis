VERSION 5.00
Begin VB.Form frmCruisePorts2 
   BackColor       =   &H00FF0000&
   Caption         =   "Alaskan Cruise Ports"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   3240
      Picture         =   "CruisePorts2.frx":0000
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   1200
      Width           =   6255
   End
   Begin VB.CommandButton cmdReturn9 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdKetchikan 
      BackColor       =   &H00404080&
      Caption         =   "Ketchikan"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmdJuneau 
      BackColor       =   &H00404080&
      Caption         =   "Juneau"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Alaskan Cruise Ports"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmCruisePorts2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmCruisePorts2
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form includes two command buttons that represent the different ports that the cruise ship
'will be stopping at throughout the trip and these buttons will bring you to one of two forms.

Option Explicit
Private Sub cmdJuneau_Click()
frmJuneau.Show
frmCruisePorts.Hide
End Sub

Private Sub cmdKetchikan_Click()
frmKetchikan.Show
frmCruisePorts.Hide
End Sub

Private Sub cmdReturn9_Click()
frmCruisePorts.Hide
frmAlaskanHome.Show
End Sub

