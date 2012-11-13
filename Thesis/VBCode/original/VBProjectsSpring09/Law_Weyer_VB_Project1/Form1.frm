VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Ticket to Ticket.txt"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Ticket"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   1800
      ScaleHeight     =   1815
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Weyer-Law European Travel Agency"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   6615
   End
   Begin VB.Menu FileSearch 
      Caption         =   "Search By..."
      Begin VB.Menu Country 
         Caption         =   "Country"
      End
      Begin VB.Menu FileBudget 
         Caption         =   "Budget"
      End
   End
   Begin VB.Menu FileCreateID 
      Caption         =   "Create ID"
      Begin VB.Menu FileID 
         Caption         =   "Name"
      End
      Begin VB.Menu FileFlightTime 
         Caption         =   "Flight Time"
      End
      Begin VB.Menu FileExchange 
         Caption         =   "Exchange Money"
      End
   End
   Begin VB.Menu FileQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()

Open App.Path & "\Ticket.txt" For Output As #2
                
Print #2, "Weyer-Law European Travel Agency"
Print #2, ID, Budget                                'Save to file for printing. We couldn't figure out
Print #2, DepartureTicketTime, ArrivalTicketTime    'how to print within the program. These are on different
Print #2, TicketTime                                'lines for readability.

Close #2

End Sub

Private Sub cmdView_Click()

picResults.Cls

picResults.Print "Weyer-Law European Travel Agency"
picResults.Print "ID Code: "; ID                                'This shows all the info entered
picResults.Print "Destination: "; Destination                   'so far and display what still
                                                                'needs to be entered because
picResults.Print Tab(0); "Departure: "; DepartureTicketTime,    'the variables are string first
picResults.Print Tab(30); "Arrival: "; ArrivalTicketTime        'defined as "insert data"
picResults.Print "**************************************************************"

If Budget > 1 Then
    picResults.Print "Total Expenses: "; FormatCurrency(Budget)
    picResults.Print "Total with Tax: "; FormatCurrency((Budget * 1.09))
    picResults.Print "Time Stamp: "; Time
End If
TicketTime = Time

End Sub


Private Sub Country_Click()
Form1.Hide          'opens form to view pictures of countries
Form3.Show
End Sub

Private Sub FileBudget_Click()
Dim BudgetTemp As Single

picResults.Cls

BudgetTemp = InputBox("Please enter your budget here.", "Budget")

BudgetTemp = Int(BudgetTemp)

Select Case BudgetTemp
    Case Is >= 2000
        picResults.Print "You can afford to go to the moon!"    'This case is to determine
    Case 1450 To 2000                                           'what countries are within
        picResults.Print "You can afford to go to Venice!"      'the user's budget.
    Case 1200 To 1449                                           'These are made up numbers.
        picResults.Print "You can afford to go to Wartburg!"
    Case 1050 To 1199
        picResults.Print "You can afford to go to Rome!"
    Case 1000 To 1049
        picResults.Print "You can afford to go to Prague!"
    Case 850 To 999
        picResults.Print "You can afford to go to Budapest!"
    Case 700 To 849
        picResults.Print "You can afford to go to Berlin!"
    Case Else
        MsgBox "Maybe you should consider paying more than $700 if you want to tavel to Europe."
End Select

End Sub


Private Sub FileExchange_Click()
Form5.Show
End Sub

Private Sub FileFlightTime_Click()
Form1.Hide
Form4.Show
End Sub

Private Sub Form_Load()

ID = "Enter an ID"                                    'Makes "error" messages, so the user knows what else
DepartureTicketTime = "Enter a Destination"           'is still required for the ticket to be full.

End Sub

Private Sub FileQuit_Click()
End
End Sub

Private Sub FileID_Click()
Form2.Show      'Shows ID form so the user can create an ID, almost like a n input box. Since all the
                'information that goes on the ticket are global strings, it stills
End Sub         'prints on this form.

