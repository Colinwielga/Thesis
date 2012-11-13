VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form4"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14745
   LinkTopic       =   "Form4"
   ScaleHeight     =   11010
   ScaleWidth      =   14745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhotel 
      Caption         =   "Continue to hotel listings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8640
      TabIndex        =   13
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdquit 
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
      Height          =   1155
      Left            =   10800
      TabIndex        =   7
      Top             =   8040
      Width           =   1995
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Choose a different resort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7440
      TabIndex        =   6
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdfindprice 
      Caption         =   "Find lift ticket price"
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
      Left            =   7440
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtseniors 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtchildren 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   3
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   7875
      TabIndex        =   2
      Top             =   7560
      Width           =   7935
   End
   Begin VB.TextBox txtyouth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtadults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   3690
      Left            =   9600
      Picture         =   "mountsnow_tickets.frx":0000
      Top             =   1920
      Width           =   4875
   End
   Begin VB.Label Label6 
      Caption         =   "Please fill in all fields. If there aren't any of the age range in your party, fill the box with a zero(0)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   13815
   End
   Begin VB.Label Label5 
      Caption         =   "Enter how many senior skiers (65+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "5 and under ski free!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Enter how many children skiers (6-12)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   5400
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Enter how many young adult skiers (13 -18)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter how many adult skiers (19-64)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   4335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Ski Trip
'Form Name: mountsnow_tickets
'Author: Sam Pilney
'Written: March 16,2009
'this page gives the lift ticket prices for the Mount Snow resort
'the user inputs how many people in each age range will need life tickets

'this subroutine brings the user back to the beginning form
Private Sub cmdback_Click()
Form4.Hide
Form1.Show
End Sub
'this subroiutine find the total price of life tickets per day depending on which resort
Private Sub cmdfindprice_Click()
Dim Adults As Integer, Children As Integer, Youth As Integer, Seniors As Integer
Dim TicketPrice As Single
picResults.Cls
Adults = txtadults.Text
Children = txtchildren.Text
Youth = txtyouth.Text
Seniors = txtseniors.Text
TicketPrice = (Adults * 75) + (Youth * 64) + (Children * 52) + (Seniors * 52)
picResults.Print "Lift tickets will cost " & FormatCurrency(TicketPrice) & " per day."
End Sub
'this subroutine brings the user to the hotels form for the corresponding resort
Private Sub cmdhotel_Click()
Form4.Hide
Form7.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub
