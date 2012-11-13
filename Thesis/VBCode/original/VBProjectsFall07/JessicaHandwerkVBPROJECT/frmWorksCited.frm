VERSION 5.00
Begin VB.Form frmWorksCited 
   BackColor       =   &H80000007&
   Caption         =   "Works Cited"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000012&
      Caption         =   "Computer Concepts and Applications for Non-Majors, by Noreen Herzfeld"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000012&
      Caption         =   "Visual Basic Tutorial - www.devdos.com/vb/lesson3.shtml"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000012&
      Caption         =   "Appending - www.garybeene.com/cod/visual%20basic13.htm"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "VB Helper - www.vb-helper.com/"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "Habitat Info - www.centralminnesotahabitat.org"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "Other Sources"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Donate Now- www.google.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "House - www.people.csail.mit.edu"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Shaking Hands- www.giftrap.com"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Requirement Diagram - www.agilmodeling.com"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Student Writing - www.istockphoto.com"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Builder - www.centralminnesotahabitat.org"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Photos"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblWorksCited 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   24
         Charset         =   0
         Weight          =   800
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   3600
      Picture         =   "frmWorksCited.frx":0000
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmWorksCited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End

End Sub

Private Sub cmdMenu_Click()
frmMenu.Show
frmWorksCited.Hide

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label5_Click()

End Sub
