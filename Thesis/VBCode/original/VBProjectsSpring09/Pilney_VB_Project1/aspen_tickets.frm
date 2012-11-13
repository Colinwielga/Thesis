VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   Caption         =   "Form2"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
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
      Top             =   6960
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
      TabIndex        =   12
      Top             =   7320
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
      Left            =   9360
      TabIndex        =   11
      Top             =   3600
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
      Left            =   9240
      TabIndex        =   10
      Top             =   1200
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
      Left            =   6240
      TabIndex        =   9
      Top             =   840
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
      Left            =   6240
      TabIndex        =   6
      Top             =   4440
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
      Left            =   360
      ScaleHeight     =   2235
      ScaleWidth      =   7875
      TabIndex        =   4
      Top             =   6840
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
      Left            =   6240
      TabIndex        =   3
      Top             =   3240
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
      Left            =   6240
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   11400
      Picture         =   "aspen_tickets.frx":0000
      Top             =   1080
      Width           =   7500
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
      Left            =   240
      TabIndex        =   14
      Top             =   240
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
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "6 and under ski free!"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Enter how many children skiers (7-12)"
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
      Left            =   840
      TabIndex        =   5
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Enter how many children skiers (13 -17)"
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
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter how many adult skiers (18-64)"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Ski Trip
'Form Name: aspen_tickets
'Author: Sam Pilney
'Written: March 16,2009
'this page gives the lift ticket prices for the Aspen resort
'the user inputs how many people in each age range will need life tickets

'this subroutine brings the user back to the beginning form
Private Sub cmdback_Click()
Form1.Show
Form2.Hide
End Sub
'this subroiutine find the total price of life tickets per day depending on which resort
Private Sub cmdfindprice_Click()
Dim Adults As Integer, Children As Integer, Youth As Integer, Seniors As Integer
picResults.Cls
Dim TicketPrice As Single
Adults = txtadults.Text
Children = txtchildren.Text
Youth = txtyouth.Text
Seniors = txtseniors.Text
TicketPrice = (Adults * 101) + (Youth * 92) + (Children * 67) + (Seniors * 92)
picResults.Print "Lift tickets will cost " & FormatCurrency(TicketPrice) & " per day."
End Sub

'this subroutine brings the user to the hotels form for the corresponding resort
Private Sub cmdhotel_Click()
Form2.Hide
Form9.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub
