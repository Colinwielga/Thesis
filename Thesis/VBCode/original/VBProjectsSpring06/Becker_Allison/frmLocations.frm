VERSION 5.00
Begin VB.Form frmLocations 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Locations"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLocations 
      BackColor       =   &H00FF8080&
      Caption         =   "List of Locations Throughout Minnesota"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin VB.PictureBox Picture2 
      Height          =   4095
      Left            =   6720
      Picture         =   "frmLocations.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   3555
      TabIndex        =   2
      Top             =   4320
      Width           =   3615
   End
   Begin VB.PictureBox picOutput2 
      Height          =   6255
      Left            =   2520
      ScaleHeight     =   6195
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF8080&
      Caption         =   "Go Back to Main Menu"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By Allison Becker"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   8160
      Width           =   5295
   End
   Begin VB.Label lblLocations 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Find a Flowers For U! Location Near You!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   9975
   End
End
Attribute VB_Name = "frmLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Flowers For U! (FlowerShop.vbp)
'Form Name: (frmLocations)
'Author: Allison Becker
'Date Written: 3/23/06
'Objective: The objective of this page is to inform the user about the locations
'that Flowers For U! has to offer.

Option Explicit

Private Sub cmdBack_Click()
    frmLocations.Hide
    frmFlowerShop.Show
End Sub

Private Sub cmdLocations_Click() 'using file output
    picOutput2.Print "Locations"
    picOutput2.Print "***************************************************"
    picOutput2.Print "Eden Prairie, MN"
    picOutput2.Print "Chaska, MN"
    picOutput2.Print "St. Cloud, MN"
    picOutput2.Print "Alexandria, MN"
    picOutput2.Print "Fridley, MN"
    picOutput2.Print "Bloomington, MN"
    picOutput2.Print "Duluth, MN"
    picOutput2.Print "Minneapolis, MN"
    picOutput2.Print "Apple Valley, MN"
    picOutput2.Print "Minnetonka, MN"
    picOutput2.Print "Grand Rapids, MN"
    picOutput2.Print "Sauk Rapids, MN"
    picOutput2.Print "Plymouth, MN"
    picOutput2.Print "St. Louis Park, MN"
    picOutput2.Print "Maple Grove, MN"
    picOutput2.Print "Eagan, MN"
    picOutput2.Print
End Sub



