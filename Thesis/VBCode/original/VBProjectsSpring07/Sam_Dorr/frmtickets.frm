VERSION 5.00
Begin VB.Form frmtickets 
   BackColor       =   &H008080FF&
   Caption         =   "Tickets"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   Picture         =   "frmtickets.frx":0000
   ScaleHeight     =   6945
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Map of Stadium"
      Height          =   975
      Left            =   9480
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      Height          =   4335
      Left            =   5280
      ScaleHeight     =   4275
      ScaleWidth      =   4635
      TabIndex        =   7
      Top             =   1320
      Width           =   4695
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmddone 
      Caption         =   "DONE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdticket4 
      BackColor       =   &H00FF8080&
      Caption         =   "$100 Ticket: Suite (Purple on Map)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton cmdtickt3 
      BackColor       =   &H00FF8080&
      Caption         =   "$500: Field level seating (Blue on Map) "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   4335
   End
   Begin VB.CommandButton cmdticket2 
      BackColor       =   &H00FF8080&
      Caption         =   "$ 25 Ticket: Main Concourse (Green Seats on Map)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton cmdticket1 
      BackColor       =   &H00FF8080&
      Caption         =   "$ 10 Ticket:  Outfield General Admission (Yellow Seats on Map)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label lbldirections 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmtickets.frx":97AA
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   9975
   End
End
Attribute VB_Name = "frmtickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'College World Series.(NCAACollegeWorldSeries.vbp)

'Form name: frmtickets; Form caption: Tickets

'Author: Sam Dorr

'Date written: March 25, 2007

' Form Objective: The objective of frmtickets is for the user to buy CWS tickets.  The
'                   range from $10 to $150.  The user also gets a chance to look at
'                   a map to make their decision easier.  The functions display, add,
'                   account for service fee and total up their costs.  There is also a
'                    select case function that displays a messeage depending on the total
'                   cost
Option Explicit

Dim Total1 As Single
Dim Total2 As Single
Dim tax As Single

Private Sub cmdclear_Click()
    Total1 = 0
    picresult.Picture = LoadPicture
End Sub

Private Sub cmdback_Click()
    frmtickets.Hide
    frmhome.Show
End Sub

Private Sub cmddone_Click()
Dim tax As Single, Total As Single
    picresults.Print "----------------------"
    picresults.Print "subtotal", Tab(40); FormatCurrency(Total1, 2) 'puts total w/o service charge
    tax = Total1 * 0.06 'finds service charge
    picresults.Print "service charge", Tab(40); FormatCurrency(tax)
    Total2 = tax + Total1 'adds service charge to total
    picresults.Print "Subtotal with service charge", Tab(40); FormatCurrency(Total2) 'prints subtotal

    Select Case Total2 'gives specific output for subtotals
        Case 0 To 49
            picresults.Print "Your total was under $50."
        Case 50 To 99
            picresults.Print "Your purchase was under $100."
        Case 100 To 199
            picresults.Print "Your purchase was under $200."
        Case 200 To 5000
            picresults.Print "Your purchase was under $200.  You will have a blast!"
    End Select
        
End Sub

Private Sub cmdreset_Click()
    Total1 = 0 'resets total
    picresults.Cls 'clears picture box
End Sub
'the follwing keeps a total while displaying the tickets in a picture box.
Private Sub cmdticket1_Click()
Dim cmdticket1 As Integer
    cmdticket1 = 10
    Total1 = Total1 + cmdticket1
    picresults.Print "Outfield General Admission ticket $10"
End Sub

Private Sub cmdticket2_Click()
Dim cmdticket2 As Single
    cmdticket2 = 25
    Total1 = Total1 + cmdticket2
    picresults.Print "Main Concourse ticket $25"
End Sub

Private Sub cmdticket4_Click()
Dim cmdticket4 As Single
    cmdticket4 = 150
    Total1 = Total1 + cmdticket4
    picresults.Print "Suite ticket $150"
End Sub

Private Sub cmdtickt3_Click()
Dim cmdticket3 As Single
    cmdticket3 = 50
    Total1 = Total1 + cmdticket3
    picresults.Print "Behind Home Plate ticket $50"
End Sub

Private Sub Command1_Click()
    frmtickets.Hide
    frmmap.Show
End Sub
