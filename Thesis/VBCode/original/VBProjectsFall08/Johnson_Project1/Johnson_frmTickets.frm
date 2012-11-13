VERSION 5.00
Begin VB.Form frmTickets 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   -1095
   ClientTop       =   -1065
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Johnson_frmTickets.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back to Home"
      Height          =   1215
      Left            =   5640
      TabIndex        =   20
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotal1 
      Caption         =   "Calculate the Total Price of Ticket(s)"
      Height          =   1215
      Left            =   5640
      TabIndex        =   19
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK TO HOME!"
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
      Left            =   6240
      TabIndex        =   17
      Top             =   12360
      Width           =   3495
   End
   Begin VB.PictureBox Picture2 
      Height          =   5535
      Left            =   8040
      Picture         =   "Johnson_frmTickets.frx":4347E2
      ScaleHeight     =   5475
      ScaleWidth      =   7155
      TabIndex        =   15
      Top             =   4920
      Width           =   7215
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Get Total        (Click Here)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1560
      TabIndex        =   14
      Top             =   11760
      Width           =   3135
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000004&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1395
      ScaleWidth      =   4995
      TabIndex        =   12
      Top             =   8760
      Width           =   5055
   End
   Begin VB.TextBox txtqty 
      Height          =   735
      Left            =   5520
      TabIndex        =   11
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtnsh 
      Height          =   615
      Left            =   5520
      TabIndex        =   9
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox txtesh 
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtSection 
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   9000
      Picture         =   "Johnson_frmTickets.frx":4B41BC
      ScaleHeight     =   4515
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000013&
      Caption         =   "Call (414) 902-4400 today to get your Brewer Tickets!!!!!!!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      TabIndex        =   18
      Top             =   13200
      Width           =   8295
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000013&
      Caption         =   "TICKET PRICE EQUATION"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000013&
      Caption         =   "Total Cost of Ticket(s):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   13
      Top             =   7680
      Width           =   5055
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000013&
      Caption         =   "How many tickets do you wish to purchase?"
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
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Note: Respond with ""yes"" or ""no"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      Caption         =   "Are you a new season seat holder?"
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
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Note: Respond with ""yes"" or ""no"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "Are you an existing season seat holder?"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   $"Johnson_frmTickets.frx":502F66
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "What kind of section do you want to sit in?"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   5295
   End
End
Attribute VB_Name = "frmTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Milwaukee Brewers Fan Club Program 2008

'Form Name: Ticket Price Equation

'Author: Matthew Johnson

'Date Written: 11/3/2008

'Objective: In this program, I construct a ticket cost calculator. It demonstrates my ability
'to set up parameters, so that I calculate something.  It shows that I can use text
'boxes as inputs.  It demonstrates my ability to use message boxes.

Private Sub cmdTotal1_Click()
'Here I declare the variables needed to construct my program
Dim section As String, existing As String, ticketPrice As Single, nsh As String, qty As Integer
Dim TotalCost As Single
picResults.Cls
'Here I set the text boxes equal to inputs, which I use when I set the parameters
'dictate ticket price
qty = txtqty.Text
section = txtSection.Text
existing = txtesh.Text
nsh = txtnsh.Text
'This section I set perameters that decide the ticket price, because depending on if you
'have season tickets or not, you have special privelages.

    If (section = "Field Diamond Box") And (existing = "yes") Then
        ticketPrice = 62
    End If
    
    If ((section = "Field Diamond Box") And (nsh = "yes")) Then
        ticketPrice = 65
    End If
    
    If section = "Field Diamond Box" And existing = "no" And nsh = "no" Then
        MsgBox "Season Seat Holders Cannot Get these Tickets...", , "Error"
    End If

    If (section = "Field Infield Box") And (existing = "yes") Then
        ticketPrice = 62
    End If
    
    If ((section = "Field Infield Box") And (nsh = "yes")) Then
        ticketPrice = 65
    End If
    
    If section = "Field Infield Box" And existing = "no" And nsh = "no" Then
        ticketPrice = 48
    End If

    If (section = "Field Outfield Box") And (existing = "yes") Then
        ticketPrice = 29
    End If
    
    If ((section = "Field Outfield Box") And (nsh = "yes")) Then
        ticketPrice = 31
    End If
    
    If section = "Field Outfield Box" And existing = "no" And nsh = "no" Then
        ticketPrice = 38
    End If

    If (section = "Loge Diamond Box") And (existing = "yes") Then
        ticketPrice = 35
    End If
    
    If ((section = "Loge Diamond Box") And (nsh = "yes")) Then
        ticketPrice = 37
    End If

    If section = "Loge Diamond Box" And existing = "no" And nsh = "no" Then
        ticketPrice = 45
    End If

    If (section = "Loge Infield Box") And (existing = "yes") Then
        ticketPrice = 29
    End If

    If ((section = "Loge Infield Box") And (nsh = "yes")) Then
        ticketPrice = 30
    End If

    If section = "Loge Infield Box" And existing = "no" And nsh = "no" Then
        ticketPrice = 36
    End If

    If (section = "Loge Outfield Box") And (existing = "yes") Then
        ticketPrice = 23
    End If

    If ((section = "Loge Outfield Box") And (nsh = "yes")) Then
        ticketPrice = 25
    End If

    If section = "Loge Outfield Box" And existing = "no" And nsh = "no" Then
        ticketPrice = 28
    End If

    If (section = "Club Infield Box") And (existing = "yes") Then
        ticketPrice = 28
    End If

    If ((section = "Club Infield Box") And (nsh = "yes")) Then
        ticketPrice = 30
    End If

    If section = "Club Infield Box" And existing = "no" And nsh = "no" Then
        ticketPrice = 40
    End If

    If (section = "Club Outfield Box") And (existing = "yes") Then
        MsgBox "Season Seat Holders don't have access to Club Outfield Box Seating", , "Error"
    End If

    If ((section = "Club Outfield Box") And (nsh = "yes")) Then
        MsgBox "Season Seat Holders don't have access to Club Outfield Box Seating", , "Error"
    End If

    If section = "Club Outfield Box" And existing = "no" And nsh = "no" Then
        ticketPrice = 36
    End If

    If (section = "Terrace Box") And (existing = "yes") Then
        ticketPrice = 15
    End If
    
    If ((section = "Terrace Box") And (nsh = "yes")) Then
        ticketPrice = 16
    End If

    If section = "Terrace Box" And existing = "no" And nsh = "no" Then
        ticketPrice = 20
    End If
    
    If (section = "Terrace Reserved") And (existing = "yes") Then
        ticketPrice = 9
    End If

    If ((section = "Terrace Reserved") And (nsh = "yes")) Then
        ticketPrice = 10
    End If

    If section = "Terrace Reserved" And existing = "no" And nsh = "no" Then
        ticketPrice = 14
    End If

    If (section = "Field Bleachers") And (existing = "yes") Then
        ticketPrice = 14
    End If

    If ((section = "Field Bleachers") And (nsh = "yes")) Then
        ticketPrice = 15
    End If

    If section = "Field Bleachers" And existing = "no" And nsh = "no" Then
        ticketPrice = 20
    End If

    If (section = "Loge Bleachers") And (existing = "yes") Then
        MsgBox "Season Seat Holders don't have access to seating in Loge Bleachers", , "Error"
    End If

    If ((section = "Loge Bleachers") And (nsh = "yes")) Then
        MsgBox "Season Seat Holders don't have access to seating in Loge Bleachers", , "Error"
    End If
    
    If section = "Loge Bleachers" And existing = "no" And nsh = "no" Then
        ticketPrice = 20
    End If

    If (section = "Bernie's Terrace") And (existing = "yes") Then
        MsgBox "Season Seat Holders don't have access to seating in Loge Bleachers", , "Error"
    End If
    
    If ((section = "Bernie's Terrace") And (nsh = "yes")) Then
        MsgBox "Season Seat Holders don't have access to seating in Loge Bleachers", , "Error"
    End If
    
    If section = "Bernie's Terrace" And existing = "no" And nsh = "no" Then
        ticketPrice = 8
    End If

'The program needs to find out whether or not the buyer is an existing season seat holder.
    If existing = "" Then
        MsgBox "You must enter no or yes where you tell  whether or not you're an existing season seat holder.", , "Error"
    End If

'The program needs to find out whether or not the buyer is a new season seat holder.
    If nsh = "" Then
        MsgBox "You must enter no or yes where you tell  whether or not you're a new season seat holder.", , "Error"
    End If

'qty must be > 0; otherwise, there would be no tickets sold
    If qty <= 0 Then
        MsgBox "You must enter a positive number in how many tickets you want to buy.", , "Error"
    End If

'the total cost of the tickets is equal to the quantity of tickets sold times the price
'of the tickets
TotalCost = (qty * ticketPrice)

'Printing off the cost of the tickets
picResults.Print "Total Ticket Price will be: "; FormatCurrency(TotalCost)

End Sub
'This button allows the program to back track to the initial page
Private Sub CmdBack_Click()
    frmIntro.Show
    frmTickets.Hide

End Sub

