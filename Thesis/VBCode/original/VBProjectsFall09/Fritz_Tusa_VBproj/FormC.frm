VERSION 5.00
Begin VB.Form LiftTicket 
   Caption         =   "Lift Ticket Cost"
   ClientHeight    =   11730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17130
   LinkTopic       =   "Form1"
   ScaleHeight     =   11730
   ScaleWidth      =   17130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FormC.frx":0000
      Top             =   8640
      Width           =   4455
   End
   Begin VB.PictureBox picCost 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   9840
      ScaleHeight     =   4875
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   1800
      Width           =   5655
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   840
         ScaleHeight     =   2295
         ScaleWidth      =   1935
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdHoMuch 
      BackColor       =   &H8000000D&
      Caption         =   "How Much will lift tickets cost you??"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "Show Resorts"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2760
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   2895
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H008080FF&
      FillColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   1320
      ScaleHeight     =   4755
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   1800
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000001&
      DataField       =   "&H00FFC0C0&"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   735
      Left            =   6000
      TabIndex        =   1
      Text            =   "Cost of a Lift Ticket!"
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdToTitleC 
      BackColor       =   &H008080FF&
      Caption         =   "To Title"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9960
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Image Image1 
      DataField       =   "&H80000005&"
      Height          =   11880
      Left            =   -360
      Picture         =   "FormC.frx":011E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17520
   End
End
Attribute VB_Name = "LiftTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SKI TRIP'
'LIFT TICKET PRICING'
'MAX TUSA'
'8-18'
'THIS FORM ASKS THE USER FOR DATA TO CALCULATE LIFT TICKET PRICES'
Option Explicit

Private Sub cmdHoMuch_Click()
'dim variables'
Dim cost As Currency, number As Integer, howmanydays As Integer
Dim rebate As Single

'set liftticket cost to zero to guard against multiple entries'
skiticketcost = 0

'clear the picture box'
picCost.Cls
picImage.Cls

'get the resort number'
place = InputBox("Which resort are you purchasing tickets to?", "Tickets")

'select case to determine cost'
Select Case place
    Case 1 To 5
        cost = 50
    Case 6 To 10
        cost = 35
    Case 11 To 13
        cost = 20
    Case Else
        MsgBox "FAIL, try again"
End Select
        
'get values for the number of days and number of people'
number = InputBox("How many tickets would you like to buy?", "How Many?")
howmanydays = InputBox("How many days are you going to need lift tickets?", "Days")
'calculate total cost'
totalCost = cost * number * howmanydays

'display results'
picCost.Print "The Cost is "; FormatCurrency(totalCost); " before rebate!"

'find amount of rebate for the number of people'
Select Case number
    Case 1 To 9
        cost = cost
    Case 10 To 19
        cost = cost - (cost * 0.25)
    Case Is > 19
        cost = cost - (cost * 0.4)
    Case Else
        picCost.Print "FAIL"
End Select

'find the amount of rebate for the number of days'
Select Case howmanydays
    Case 1 To 4
        cost = cost
    Case 4 To 7
        cost = cost - (cost * 0.15)
    Case Is > 7
        cost = cost - (cost * 0.25)
    Case Else
        picCost.Print "FAIL"
End Select

'calculate the new price'
rebate = cost * number * howmanydays

picCost.Print ""
'display new results'
picCost.Print "Your price with rebates is "; FormatCurrency(rebate)

picImage.Picture = LoadPicture(App.Path & "\images\" & "\instructor.jpg")

'add cost to total cost'
skiticketcost = rebate
    
End Sub

Private Sub cmdToTitleC_Click()
Title.Show
LiftTicket.Hide
End Sub

Private Sub Command1_Click()
'dim variables'
Dim z As Integer, pass As Integer, pos As Integer, tempResort As String, tempRuns As Single

'clear the picture box'
picResults2.Cls

'display resorts sorted from biggest to smallest and number associated with it'
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If skiruns(pos) < skiruns(pos + 1) Then
            tempResort = resorts(pos)
            tempRuns = skiruns(pos)
            resorts(pos) = resorts(pos + 1)
            skiruns(pos) = skiruns(pos + 1)
            resorts(pos + 1) = tempResort
            skiruns(pos + 1) = tempRuns
            End If
    Next pos
Next pass

'pritn header'
picResults2.Print "#", "Resort Name", , "Number of Runs"
'print results'
For z = 1 To ctr
    picResults2.Print z; ")"; Tab(15); resorts(z); Tab(47); skiruns(z)
Next z

End Sub

