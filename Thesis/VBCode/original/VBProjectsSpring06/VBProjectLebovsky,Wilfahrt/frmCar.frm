VERSION 5.00
Begin VB.Form frmCar 
   BackColor       =   &H00000000&
   Caption         =   "Picking your Car"
   ClientHeight    =   7440
   ClientLeft      =   2835
   ClientTop       =   1725
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H80000009&
      Caption         =   "Previous Page"
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdCar4 
      BackColor       =   &H00FF8080&
      Caption         =   "Pick Me"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdCar3 
      BackColor       =   &H00808080&
      Caption         =   "Pick Me"
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdCar2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pick Me"
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdCar1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pick Me"
      Height          =   735
      Left            =   7200
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox picBMW 
      Height          =   975
      Left            =   6480
      Picture         =   "frmCar.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.PictureBox picVan 
      Height          =   2415
      Left            =   6600
      Picture         =   "frmCar.frx":126B
      ScaleHeight     =   2355
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   4560
      Width           =   2295
   End
   Begin VB.PictureBox picBuggy 
      Height          =   1695
      Left            =   600
      Picture         =   "frmCar.frx":38D6
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   5040
      Width           =   2415
   End
   Begin VB.PictureBox picClassic 
      Height          =   975
      Left            =   600
      Picture         =   "frmCar.frx":71CA
      ScaleHeight     =   915
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Clay Wilfahrt and Andy Lebovsky"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Click on car to see statistics."
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1395
      Left            =   3000
      TabIndex        =   9
      Top             =   2760
      Width           =   3450
   End
End
Attribute VB_Name = "frmCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frmCar(frmCar.frm)
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'The Objective of this form is for the user to gain information about the cars and to pick one of them to race.
Option Explicit
'Brings you to Race screen 1
Private Sub cmdCar1_Click()
    frmRace1.Show
    frmCar.Hide
End Sub
'Brings you to Race screen 2
Private Sub cmdCar2_Click()

    frmRace2.Show
    frmCar.Hide
End Sub
'Brings you to Race screen 3
Private Sub cmdCar3_Click()

    frmRace3.Show
    frmCar.Hide
End Sub
'Brings you to Race screen 4
Private Sub cmdCar4_Click()

    frmRace4.Show
    frmCar.Hide
End Sub
'Brings you to Main screen
Private Sub cmdPrevious_Click()
    frmmain.Show
    frmCar.Hide
End Sub
Private Sub Form_Load()
    frmmain.Show
    frmCar.Hide
End Sub
'Gives info about the BMW
Private Sub PicBMW_Click()
    MsgBox Left("This BMW is made from out of raw power.  It tops out at a speed of 120 mph and has the gas mileage of 30 mpg.  The driver should feel fortunate to drive it in today's race.", 172), , "BMW"
End Sub
'Gives info about the Classic Antique Car
Private Sub PicClassic_Click()
    MsgBox Right("The Classic Antique Car, made in 1934, is built like a tank.  Made to sustain bullet shots, it is pure power.  The maximum speed is 60 mph and the gas mileage is 15 mpg.", 170), , "Classic Antique Car"
End Sub
'Gives info about the Buggy
Private Sub PicBuggy_Click()
    MsgBox "Oh boy, the Buggy made out of wood is a pure joy to drive.  With a top speed of 20 mph, a gas mileage of 8 mpg, and the chance of a flat tire every mile, it's a blast.", , "Buggy"
End Sub
'Gives info about the Mini-Van
Private Sub PicVan_Click()
    MsgBox "The Turbo Mini-Van is perfect for hauling hyper children around the town.  With a top speed of 2 mph, you know it's safe.  Also, with a gas mileage of 0 mpg, you can drive for hours.  Trust me it's sweet.", , "Turbo Mini-Van"
End Sub


