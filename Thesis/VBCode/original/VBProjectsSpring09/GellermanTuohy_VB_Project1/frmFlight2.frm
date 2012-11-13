VERSION 5.00
Begin VB.Form frmFlight 
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16590
   LinkTopic       =   "Form1"
   ScaleHeight     =   11040
   ScaleWidth      =   16590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFlight 
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdMatchAndStopSearch 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   7
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox txtTickets 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextActivityPage 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   4
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton cmdFlightOptions 
      Caption         =   "View Available Flights"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   5535
   End
   Begin VB.PictureBox picResults 
      Height          =   4335
      Left            =   7560
      ScaleHeight     =   4275
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   2280
      Width           =   6975
   End
   Begin VB.Label Label3 
      Caption         =   "Please enter flight ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   8
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "No. of tickets:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Gotta Catch A Flight!!!!!!!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   9495
   End
End
Attribute VB_Name = "frmFlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tickets As Integer
Dim FlightID(1 To 50) As Integer, TakeOff(1 To 50) As String, TicketCost(1 To 50) As Single
Dim Temp As Integer, Temp1 As String, Temp2 As Integer
Dim FlightNum As Integer
Dim GoNext As Boolean

Private Sub cmdFlightOptions_Click()
    Open App.Path & "\Flights.txt" For Input As #1
    
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, FlightID(CTR), TakeOff(CTR), TicketCost(CTR)
    Loop
    Close #1
    
    picResults.Print "Flight ID Number"; Tab(22); "Flight Departure to Destination"; Tab(55); "Ticket Cost Per Person"
    picResults.Print "************************************************************************************************************"
    
    For I = 1 To CTR
        picResults.Print FlightID(I); Tab(22); TakeOff(I); Tab(55); FormatCurrency(TicketCost(I))
    Next I
    
    
    
End Sub

Private Sub cmdMatchAndStopSearch_Click()
    
    Found = False
    I = 0
    
    FlightNum = CInt(txtFlight.Text)
    
    'this part checks for st joe
    
    If Location = "St. Joe" Then
        Do While ((Not Found) And (I < CTR))    'searches for flight ID number and then once it has found an ID number,
            I = I + 1                           'it saves the flight ID number as I and Found becomes True
            If FlightNum = FlightID(I) And FlightNum >= 1 And FlightNum <= 4 Then
                Found = True
            End If
        Loop
        
         If (Not Found) Then
            MsgBox "Flight " & FlightNum & " is not a valid flight to " & Location & " Please Try Again", vbExclamation
            txtFlight.Text = ""
            
        Else
            MsgBox "You have selected flight number " & FlightID(I) & " which is the flight from " & TakeOff(I) & "."
            GoNext = True
        End If
    End If
    
    'this part checks for Normandy
    
    If Location = "Normandy" Then
        Do While ((Not Found) And (I < CTR))    'searches for flight ID number and then once it has found an ID number,
            I = I + 1                           'it saves the flight ID number as I and Found becomes True
            If FlightNum = FlightID(I) And FlightNum >= 5 And FlightNum <= 8 Then
                Found = True
            End If
        Loop
        
         If (Not Found) Then
            MsgBox "Flight " & FlightNum & " is not a valid flight to " & Location & " Please Try Again", vbExclamation
            txtFlight.Text = ""
            
        Else
            MsgBox "You have selected flight number " & FlightID(I) & " which is the flight from " & TakeOff(I) & "."
            GoNext = True
        End If
    End If
    
    'this part checks for canada
    
    If Location = "Saskatchewan" Then
        Do While ((Not Found) And (I < CTR))    'searches for flight ID number and then once it has found an ID number,
            I = I + 1                           'it saves the flight ID number as I and Found becomes True
            If FlightNum = FlightID(I) And FlightNum >= 9 And FlightNum <= 12 Then
                Found = True
            End If
        Loop
        
         If (Not Found) Then
            MsgBox "Flight " & FlightNum & " is not a valid flight to " & Location & " Please Try Again", vbExclamation
            txtFlight.Text = ""
            
        Else
            MsgBox "You have selected flight number " & FlightID(I) & " which is the flight from " & TakeOff(I) & "."
            GoNext = True
        End If
    End If
    
    'this part checks for south dakota
    
    If Location = "Badlands" Then
        Do While ((Not Found) And (I < CTR))    'searches for flight ID number and then once it has found an ID number,
            I = I + 1                           'it saves the flight ID number as I and Found becomes True
            If FlightNum = FlightID(I) And FlightNum >= 13 And FlightNum <= 16 Then
                Found = True
            End If
        Loop
        
         If (Not Found) Then
            MsgBox "Flight " & FlightNum & " is not a valid flight to " & Location & " Please Try Again", vbExclamation
            txtFlight.Text = ""
            
        Else
            MsgBox "You have selected flight number " & FlightID(I) & " which is the flight from " & TakeOff(I) & "."
            GoNext = True
        End If
    End If
    
End Sub

Private Sub cmdNextActivityPage_Click()
 
 If Location = "St. Joe" And GoNext = True Then
        frmFlight.Hide
        frmActivitiesJosepth.Show
    ElseIf Location = "Normandy" Then
        frmFlight.Hide
        frmActivitiesNormandy.Show
    ElseIf Location = "Badlands" Then
        frmFlight.Hide
        frmBadlandsActivity.Show
    ElseIf Location = "Saskatchewan" Then
        frmFlight.Hide
        frmCanada.Show
Else
MsgBox "Your flight does not end at your requested vacationing spot", vbExclamation
End If
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Form_Load()
GoNext = False
End Sub
