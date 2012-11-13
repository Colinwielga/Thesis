VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9480
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmbCountDelivery 
      Caption         =   "Count Deliveries"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   8
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdSales 
      Caption         =   "Calculate Sales"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdExistingCustomer 
      BackColor       =   &H008080FF&
      Caption         =   "Existing Customer"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      MaskColor       =   &H008080FF&
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdNewCustomer 
      Caption         =   "New Customer"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   1320
      Picture         =   "frmWelcome.frx":0000
      ScaleHeight     =   1785
      ScaleWidth      =   6165
      TabIndex        =   0
      Top             =   480
      Width           =   6165
   End
   Begin VB.Label lblAdmin 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Administration Tasks"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5520
      TabIndex        =   7
      Top             =   3120
      Width           =   2820
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Today's Date:"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1110
   End
   Begin VB.Label lblDateNow 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label lblEnter 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter a new order"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCountDelivery_Click()
Dim OrderDate(1 To 1000) As String, Total(1 To 1000) As String, DeliveryDate(1 To 1000) As String
Dim TelNo(1 To 1000) As String, Ctr As Integer, FirstName(1 To 1000) As String, LastName(1 To 1000) As String
'Read file
Open App.Path & "\TheLaundryCo.txt" For Input As #1
'Initialize variables
Ctr = 0
Do While Not EOF(1)  'this loop reads data from a file into three arrays
    Ctr = Ctr + 1   'increment the Ctr
    Input #1, FirstName(Ctr), LastName(Ctr), TelNo(Ctr), OrderDate(Ctr), Total(Ctr), DeliveryDate(Ctr)
Loop
Close #1

Dim I As Integer, myDate As String, NumberOfDeliveries As Single, Found As Boolean
NumberOfDeliveries = 0 'initializing NumberOfDeliveries
Found = False
myDate = InputBox("Please enter delivery date(mm/dd/yyyy).")
I = 0

If myDate <> "" Then 'Checks to see if the user has entered date
    Do While (I < Ctr)  'searches until end of file"
        I = I + 1
        If myDate = Trim$(DeliveryDate(I)) Then 'Removes both leading and trailing blank spaces from a string
            Found = True
            NumberOfDeliveries = NumberOfDeliveries + 1
        End If
    Loop


    If Found = False Then
        MsgBox ("Take a break! There are no deliveries for " & myDate)
    
    Else: MsgBox ("You have to make " & NumberOfDeliveries & " delivery(s) for " & myDate)
    End If

Else: MsgBox ("Please enter a date to search.")
End If

End Sub

Private Sub cmdExistingCustomer_Click()
Dim TelNo(1 To 1000) As String, Ctr As Integer, FirstName(1 To 1000) As String, LastName(1 To 1000) As String


'Read file
Open App.Path & "\TheLaundryCo.txt" For Input As #1
'Initialize variables
Ctr = 0
Do While Not EOF(1)  'this loop reads data from a file into three arrays
    Ctr = Ctr + 1   'increment the Ctr
    Input #1, FirstName(Ctr), LastName(Ctr), TelNo(Ctr)
Loop
Close #1

Dim Found As Boolean, I As Integer, Number As String
Number = InputBox("Please enter customer's telephone number:")
I = 0
Found = False

If Number <> "" Then

    Do While (Not Found) And (I < Ctr)  'searches until found of end of file"
         I = I + 1
        If Number = TelNo(I) Then
            Found = True
            'If customer is found global variables are updates with the customer data
            CustFirstName = FirstName(I)
            CustLastName = LastName(I)
            CustTelNo = TelNo(I)
            MsgBox ("The following customer record was found: " & FirstName(I) & ", " & LastName(I) & ", " & TelNo(I))
        
    'close form
    Form1.Show
    frmWelcome.Hide
        End If
    Loop
    
    If (Not Found) Then
        MsgBox ("The customer was not found.")
    End If
    
Else: MsgBox ("You must enter a number to search")
End If

End Sub


Private Sub cmdNewCustomer_Click()
Form1.Show
frmWelcome.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSales_Click()
Dim OrderDate(1 To 1000) As String, Total(1 To 1000) As String, DeliveryDate(1 To 1000) As String
Dim TelNo(1 To 1000) As String, Ctr As Integer, FirstName(1 To 1000) As String, LastName(1 To 1000) As String
'Read file
Open App.Path & "\TheLaundryCo.txt" For Input As #1
'Initialize variables
Ctr = 0
Do While Not EOF(1)  'this loop reads data from a file into three arrays
    Ctr = Ctr + 1   'increment the Ctr
    Input #1, FirstName(Ctr), LastName(Ctr), TelNo(Ctr), OrderDate(Ctr), Total(Ctr), DeliveryDate(Ctr)
Loop
Close #1

Dim I As Integer, myDate As String, TotalSales As Single, Found As Boolean
TotalSales = 0 'initializing TotalSales
Found = False
myDate = InputBox("Please enter sales date(mm/dd/yyyy).")
I = 0
Do While (I < Ctr)  'searches until found of end of file"
    I = I + 1
    If myDate = Trim$(OrderDate(I)) Then 'Removes both leading and trailing blank spaces from a string
        Found = True
        TotalSales = TotalSales + Total(I)
    End If

Loop


If Found = False Then
    MsgBox ("There were no sales on the date entered.")
    
Else: MsgBox ("The total sales for " & OrderDate(I) & " was " & FormatCurrency(TotalSales))
End If


End Sub

Private Sub Form_Load()
lblDateNow.Caption = DateValue(Now)
End Sub

Private Sub Label1_Click()

End Sub
