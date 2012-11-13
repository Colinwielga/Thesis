VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00800000&
   Caption         =   "About the Program"
   ClientHeight    =   3090
   ClientLeft      =   5370
   ClientTop       =   4110
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H0000FFFF&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Money Manager 2005 (ProjectNataliesMoneyPlanner)
'frmAbout (frmAbout.frm)
'by Natalie Bly
'10/29/05
'This form allows the user to view the general purpose of the program.

Option Explicit             'makes the code easier to debug
Private Sub cmdDone_Click()
    frmAbout.Hide           'brings the user back to the Menu screen
    frmMenu.Show
End Sub

