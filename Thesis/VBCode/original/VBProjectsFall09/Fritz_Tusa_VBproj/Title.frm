VERSION 5.00
Begin VB.Form Title 
   Caption         =   "SKI TRIP!"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   Picture         =   "Title.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton gotoGrandTotal 
      BackColor       =   &H0000FF00&
      Caption         =   "Find out your GRAND TOTAL!!!"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton cmdSwitchtoFormF 
      BackColor       =   &H00008000&
      Caption         =   "What level ski runs to go down"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdSwitchtoFormE 
      BackColor       =   &H00008000&
      Caption         =   "Hotel Cost"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdSwitchtoFormD 
      BackColor       =   &H00008000&
      Caption         =   "Ski Rental Cost"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdSwitchtoFormC 
      BackColor       =   &H00008000&
      Caption         =   "Lift Ticket Cost"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdSwitchtoFormB 
      BackColor       =   &H00008000&
      Caption         =   "Possible Ski Resorts"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdSwitchform1 
      BackColor       =   &H00008000&
      Caption         =   "Cost of Airfare"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label trip 
      BackColor       =   &H80000013&
      Caption         =   "SKI TRIP TO COLORADO!"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   8895
   End
   Begin VB.Image Image1 
      Height          =   8160
      Left            =   0
      Picture         =   "Title.frx":AFCC2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11880
   End
End
Attribute VB_Name = "Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SKI TRIP'
'TITLE'
'MAX TUSA'
'8-18'
'THIS FORM IS THE MAIN SCREEN FOR THE PROJECT'

Option Explicit



        
'quit button'
Private Sub cmdQuit_Click()
End
End Sub

'go to the airfare form'
Private Sub cmdSwitchform1_Click()
Title.Hide
Airfare.Show
End Sub

'go to the ski resort form'
Private Sub cmdSwitchtoFormB_Click()
Title.Hide
SkiResorts.Show
End Sub

'go to the lift ticket form'
Private Sub cmdSwitchtoFormC_Click()
Title.Hide
LiftTicket.Show
End Sub

'go to the ski rental form'
Private Sub cmdSwitchtoFormD_Click()
Title.Hide
FormRent.Show
End Sub

'go to the hotel form'
Private Sub cmdSwitchtoFormE_Click()
Title.Hide
FormHotel.Show
End Sub

'go to the ski runs form'
Private Sub cmdSwitchtoFormF_Click()
Title.Hide
FormRUNS.Show
End Sub


'go to the grand total button'
Private Sub gotoGrandTotal_Click()
Title.Hide
GrandTotal.Show
End Sub

