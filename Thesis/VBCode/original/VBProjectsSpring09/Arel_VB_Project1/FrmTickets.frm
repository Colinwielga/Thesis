VERSION 5.00
Begin VB.Form FrmTickets 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Back To Home Schedule"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox txtbuyquantity 
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   8280
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Order Tickets!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox txtbuydate 
      BackColor       =   &H00FFC0FF&
      Height          =   405
      Left            =   8280
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtbuysection 
      BackColor       =   &H00FFC0C0&
      Height          =   405
      Left            =   8280
      TabIndex        =   6
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtbuymonth 
      BackColor       =   &H00C0FFC0&
      Height          =   405
      Left            =   8280
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtbuygame 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   8280
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Date of Game"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Month of Game"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Opponent"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7920
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   240
      Picture         =   "FrmTickets.frx":0000
      Top             =   480
      Width           =   7500
   End
End
Attribute VB_Name = "FrmTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmTickets
'Project By: Stephanie Arel
'Date Written: 3/16/2009
'The purpose of this form is to let the user purchase tickets.
Option Explicit


Private Sub Command1_Click()
'Ends the program

End
End Sub

Private Sub Command2_Click()
'Purchases tickets for the user.
Dim section(1 To 75) As Integer
Dim Price(1 To 75) As Integer
Dim Ctr As Integer

Dim buygame As String
Dim buymonth As String
Dim buydate As Single
Dim buysection As Integer
Dim buyquantity As Single
Dim Found As Boolean
Dim I As Integer
Dim buyctr As Integer
Dim Total As Single

Open App.Path & "\Ticketing.Txt" For Input As #2
Do While Not EOF(2)
    Ctr = Ctr + 1
    Input #2, section(Ctr), Price(Ctr)
Loop

Found = False
I = 0
buygame = txtbuygame.Text
buymonth = txtbuymonth.Text
buydate = txtbuydate.Text
buysection = txtbuysection.Text
buyquantity = txtbuyquantity.Text

Do While Not Found And I < Ctr
    I = I + 1
    If buysection = section(I) Then
        Total = buyquantity * Price(I)
        Found = True
        MsgBox "Thank you for purchasing tickets to the " & buygame & " versus the Twins on " & buymonth & " " & buydate & ".", , "Thank You!"
        MsgBox "You have chosen section " & buysection & ". The total cost for " & buyquantity & " tickets is " & FormatCurrency(Total, 2) & ".", , "Thank You!"
    End If
Loop

If Not Found Then
    MsgBox "I'm sorry! You have selected seats that do not exist. Please start over!", , "Error!"
End If

    End Sub

Private Sub Command3_Click()
'Takes the user back to the main menu
FrmTickets.Hide
FrmMain.Show
End Sub

Private Sub Command4_Click()
'Takes the user back to the schedule viewer.
FrmTickets.Hide
FrmSchedule.Show

End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
