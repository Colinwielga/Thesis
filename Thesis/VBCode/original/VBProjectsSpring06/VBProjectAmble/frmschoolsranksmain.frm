VERSION 5.00
Begin VB.Form frmschoolsranksmain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6360
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdaustinregion 
      Caption         =   "Austin"
      Height          =   735
      Left            =   5760
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdsyracuseregion 
      Caption         =   "Syracuse"
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdalbuquerqueregion 
      Caption         =   "Albuquerque"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdchicagoregion 
      Caption         =   "Chicago"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Image imgncaa 
      Height          =   3000
      Left            =   2760
      Picture         =   "frmschoolsranksmain.frx":0000
      Top             =   1320
      Width           =   2445
   End
   Begin VB.Label lblheader 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on one of the Regions to View"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "frmschoolsranksmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form gives the user the option to choose one of the four'
'regions, so they can view the teams and ranks.'
Option Explicit
'This button brings the user to albuquergue region form.'
Private Sub cmdalbuquerqueregion_Click()
    frmschoolsranksmain.Visible = False
    frmalbuquerqueregion.Visible = True
End Sub
'This button brings the user to austin region form.'
Private Sub cmdaustinregion_Click()
    frmschoolsranksmain.Visible = False
    frmaustinregion.Visible = True
End Sub
'This button enables the user to go back to the main form'
Private Sub cmdback_Click()
    frmschoolsranksmain.Visible = False
    frmmain.Visible = True
End Sub
'This button brings the user to chicago region form.'
Private Sub cmdchicagoregion_Click()
    frmschoolsranksmain.Visible = False
    frmchicagoregion.Visible = True
End Sub
'This button enables the user to exit the program'
Private Sub cmdexit_Click()
    End
End Sub
'This button brings the user to syracuse region form.'
Private Sub cmdsyracuseregion_Click()
    frmschoolsranksmain.Visible = False
    frmsyracuseregion.Visible = True
End Sub

