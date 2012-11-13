VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Main"
   ClientHeight    =   9855
   ClientLeft      =   645
   ClientTop       =   345
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   11745
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H008080FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   1080
      TabIndex        =   5
      Top             =   8880
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   1800
      Picture         =   "Dance Project.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   3135
      Left            =   6000
      Picture         =   "Dance Project.frx":3A0C
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   3960
      Picture         =   "Dance Project.frx":6CA3
      ScaleHeight     =   2355
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   6120
      Width           =   3255
   End
   Begin VB.CommandButton cmdRegister 
      BackColor       =   &H008080FF&
      Caption         =   "REGISTER FOR DANCE CLASSES!!!!!"
      Height          =   1335
      Left            =   6600
      TabIndex        =   1
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H008080FF&
      Caption         =   "SHOES, LEOTARDS, DRESSES, SKIRTS, UNITARDS, SOCKS, AND ACCESSORIES!"
      Height          =   1335
      Left            =   1920
      TabIndex        =   0
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   9360
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmMain (Main.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to let the user choose between buying items related to dance
                    'or to register for different classes

Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.

Private Sub cmdBuy_Click()
frmShoesetc.Show
frmMain.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRegister_Click()
    frmRegistration.Show
    frmMain.Hide
End Sub

Private Sub Form_Load()
frmShoesetc.Show
End Sub
