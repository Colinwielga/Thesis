VERSION 5.00
Begin VB.Form frmBAC 
   BackColor       =   &H00C00000&
   Caption         =   "Calculate BAC"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18780
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   18780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H0080FF80&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FF80&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   2055
   End
   Begin VB.PictureBox picResults2 
      Height          =   5775
      Left            =   10800
      Picture         =   "BAC.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   7515
      TabIndex        =   3
      Top             =   120
      Width           =   7575
   End
   Begin VB.PictureBox picResults 
      Height          =   5775
      Left            =   3120
      Picture         =   "BAC.frx":10B7E
      ScaleHeight     =   5715
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   120
      Width           =   7575
      Begin VB.PictureBox Picture1 
         Height          =   4575
         Left            =   7560
         ScaleHeight     =   4575
         ScaleWidth      =   15
         TabIndex        =   2
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.CommandButton cmdCalculateBAC 
      BackColor       =   &H0080FF80&
      Caption         =   "Click Here to Calculate Blood Alcohol Content"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label lblChart 
      BackColor       =   &H0080FF80&
      Caption         =   "Determine Your Blood Alcohol Percentage by looking at the chart"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "frmBAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Project Name: The Beer Experience
'Author: Lauren Gooley and Tim Janssen
'Date: 3-19-08
'This form calculates the blood alcohol content of a person. Using input boxes, the user enters
'their blood alcohol percentage found on the chart and the number of hours since their 1st drink.
'This form then muliplies .015(the rate your body breaks down alcohol) times the number of hours
'since your 1st drink; this number is then subracted from the blood alcohol percentage to determine the BAC
'of the user, which is given as a message box.

Private Sub cmdCalculateBAC_Click()
Dim BAC As Single, Hours As Integer, BACPercent As Single
BACPercent = InputBox("Enter your Blood Alcohol Percentage from the table")
Hours = InputBox("Enter the number of hours since your first drink")
BAC = BACPercent - (0.015 * Hours)
MsgBox ("Your Blood Alcohol Content is " & FormatNumber(BAC, 3))
End Sub

Private Sub cmdGoBack_Click()
frmBAC.Hide
frmStartUp.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
