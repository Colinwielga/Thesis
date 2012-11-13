VERSION 5.00
Begin VB.Form frmNew1 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   4245
   ClientTop       =   3030
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00808080&
      Caption         =   "Back to Main Page"
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1695
   End
   Begin VB.PictureBox picCars 
      Height          =   1575
      Left            =   3480
      Picture         =   "frmNew1.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.PictureBox picTrucks 
      Height          =   1575
      Left            =   3480
      Picture         =   "frmNew1.frx":10B2
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox picSUV 
      Height          =   1575
      Left            =   3480
      Picture         =   "frmNew1.frx":2042
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdCars 
      Caption         =   "Go to Cars"
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdTrucks 
      Caption         =   "Go to Trucks"
      Height          =   1575
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdSUV 
      BackColor       =   &H000000C0&
      Caption         =   "Go to SUVs"
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmNew1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name- Stack's Car Lot
'Form Name- frmNew1
'Author- Nick Stackenwalt
'Date Written- Saturday March 08, 2008
'Objective- This page will allow the user to choose what type of vehicle they want
'Other comments- The SUV button will take them to our New 2008 SUV models
                'The Cars button will take them to our New 2008 Car models
                'The Truck button will take them to our New 2008 Truck models
                
Private Sub cmdCars_Click()
frmNew1.Hide        'Hides New Vehicle form
frmCars.Show        'Shows Cars form
End Sub

Private Sub cmdMain_Click()
frmNew1.Hide
frmMain.Show
End Sub

Private Sub cmdSUV_Click()
frmNew1.Hide        'Hides New Vehicle form
frmSUV.Show         'Shows SUV form
End Sub

Private Sub cmdTrucks_Click()
frmNew1.Hide        'Hides New Vehicle form
frmTruck.Show       'Shows Truck form
End Sub
