VERSION 5.00
Begin VB.Form frmDNR 
   BackColor       =   &H80000001&
   Caption         =   "Minnesota DNR Project"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   5400
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   7215
      Left            =   3240
      Picture         =   "DNR.frx":0000
      ScaleHeight     =   7155
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   600
      Width           =   4815
   End
   Begin VB.CommandButton cmdTrip 
      BackColor       =   &H0000FF00&
      Caption         =   "Build Your Own Trip"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdFishing 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go to Fishing Page"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdHunt 
      BackColor       =   &H000080FF&
      Caption         =   "Load Hunting Prices and Go to Hunting Page"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Andrew Forliti             Casey Orthaus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000001&
      Caption         =   "Minnesota Outdoors And Building Your Own Trip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmDNR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Outdoors
'DNR
'Andrew Forliti and Casey Orthaus
'October 19th, 2009
'this is the main page that lets the user select which form to go into

Private Sub cmdFishing_Click()

frmDNR.Hide
frmFishing.Show

End Sub

Public Sub cmdHunt_Click()

HuntingCtr = 0   'setting the counter to 0

Open App.Path & "\Resident.txt" For Input As #1  'loading the text file

'Reading the file
Do While Not EOF(1)
    HuntingCtr = HuntingCtr + 1
    Input #1, Animal(HuntingCtr), HuntingPrice(HuntingCtr)
Loop
  
Close #1   'closing the file

'go to hunting form

frmDNR.Hide
frmHunting.Show


End Sub

Private Sub cmdQuit_Click()
End
End Sub



Private Sub cmdTrip_Click()

frmTrip.Show
frmDNR.Hide

End Sub

