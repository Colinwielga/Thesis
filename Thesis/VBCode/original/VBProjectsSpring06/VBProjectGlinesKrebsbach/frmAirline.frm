VERSION 5.00
Begin VB.Form frmAirline 
   Caption         =   "Airfare"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   BeginProperty Font 
      Name            =   "Rockwell"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmAirline.frx":0000
   ScaleHeight     =   9285
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Resorts"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   9240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOther 
      BackColor       =   &H00C0C000&
      Caption         =   "Things to consider when traveling to either Denver or Vail/Eagle Airports"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   4815
   End
   Begin VB.CommandButton cmdEagle 
      Caption         =   "Minneapolis/St. Paul International to Eagle County Regional Airport"
      Height          =   855
      Left            =   12120
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdDenver 
      Caption         =   "Minneapolis/St. Paul International to Denver International"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   10680
      Width           =   2775
   End
   Begin VB.Label lblpick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pick your Destination!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   8535
   End
End
Attribute VB_Name = "frmAirline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmAirline(frmAirline.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the suser to look up the cheapest flights from
'Minneapolis,MN to Colorado. we chose the three main airlines from Msp,MN to Denver,Colorado and
'researched the lowest prices for a one week stay during CSB/SJU's respective spring
'break dates. we also researched the lowest flight available from MN to Eagle Regional
'Airport.this form also allows the user to research insider tips on which airport to arrive
'at

Private Sub cmdDenver_Click()
    frmAirline.Hide
    frmDIA.Show

End Sub

Private Sub cmdback_Click()
    frmAirline.Hide
    frmContents.Show
End Sub

Private Sub cmdEagle_Click()
    frmAirline.Hide
    frmEagle.Show
End Sub

Private Sub cmdOther_Click()
    MsgBox "If you fly into Eagle/Vail you can save money by not renting a car and taking frequent Resort shuttles.  There is also a shuttle from the Eagle/Vail Airport to your destination.  If you plan on fliying into Denver International many spend extra money to rent a car however there are less frequent shuttles to your destination.  Another downside to flying into Denver is that it is at least an hour out from the nearest ski resorts.  Sometimes it is nice to have your own rental vehicle so that you can see what else Colorado has to give.  If you are there just to ski/snowboard a rental car is not necassary but if you are traveling alot a rental car wouldn't be a bad idea!", , "Take into Consideration"
End Sub

Private Sub Command1_Click()
frmAirline.Hide
frmForm.Show

End Sub
