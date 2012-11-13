VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   4455
   ClientTop       =   1635
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   6645
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000C&
      Caption         =   "Quit"
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   855
   End
   Begin VB.PictureBox picMustang 
      Height          =   1455
      Left            =   2280
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdUsed 
      BackColor       =   &H00000080&
      Caption         =   "Enter Used Cars"
      Height          =   1575
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H000000C0&
      Caption         =   "Enter New Cars"
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lblCars2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Cars"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblCars 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "New and Used"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblStacks 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Stack's Lot"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name- Stack's Car Lot
'Form Name- frmMain
'Author- Nick Stackenwalt
'Date Written- Saturday March 08, 2008
'Objective- This form introduces the user to our lot
            'The user can then choose to either view our New inventory or our Used inventory
'Other comments- The New button takes the user to a page where they can view our new SUVs, Trucks, or Cars
                'The Used button takes the user to a page where tehy can view our used vehicles.
                
Private Sub cmdNew_Click()
frmMain.Hide        'Hides Main form
frmNew1.Show        'Shows New Vehicles form
End Sub

Private Sub cmdQuit_Click()
End     'Ends Program
End Sub

Private Sub cmdUsed_Click()
frmMain.Hide        'Hides Main form
frmUsed1.Show       'Shows Used Cars form
End Sub

