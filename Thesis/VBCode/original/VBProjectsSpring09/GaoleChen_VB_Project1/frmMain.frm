VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort in price order"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdSelf 
      Caption         =   "No more than $?"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdUnder20 
      Caption         =   "Under $20"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   7560
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   4935
      Left            =   5760
      ScaleHeight     =   4875
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "All dishes"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdMain4 
      Caption         =   "Tender Pork-$20"
      Height          =   1815
      Left            =   2640
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdMain3 
      Caption         =   "Kongbao Chicken-$18"
      Height          =   1815
      Left            =   120
      Picture         =   "frmMain.frx":12F6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2115
   End
   Begin VB.CommandButton cmdMain2 
      Caption         =   "Honey Meatballs-$25"
      Height          =   1815
      Left            =   2760
      Picture         =   "frmMain.frx":215D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdMain1 
      Caption         =   "Prawns-$20"
      Height          =   1695
      Left            =   120
      Picture         =   "frmMain.frx":30DA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Caption         =   "<==Display in price ascending order"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   19
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "<==Ready to order? Click here!"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   17
      Top             =   7440
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "<==How much would you like to pay?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   6240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "<==More dishes under $20"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "<==Want some more?Click here!"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label lblMain 
      BackColor       =   &H000080FF&
      Caption         =   "Click the buttons to taste our most popular dishes! We will also cook what you want from the complement menu(-;"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Digital Menu
'Form Name: frmMain
'Authors: Gaole Chen
'Date Written: 3/7/09
'Objective: The user can order the dishes by clicking the vivid pictures.
'The user also may choose other dishes by opening the file.
'The form also calculate the total price automatically for the user.

Option Explicit
Dim runningTotal As Single, Dishes(1 To 30) As String, Price(1 To 30) As Integer, CTR As Integer

Private Sub cmdBack_Click()
frmMain.Hide
frmSalad.Show
End Sub

Private Sub cmdClear_Click()
picResults.Cls
End Sub

Private Sub cmdMain1_Click()
'declare the variables
Dim Numberone As Integer, Totalone As Integer
'calculate how many Lotus the user wants
Numberone = InputBox("How many Prawns would you like?")
Totalone = 20 * Numberone
runningTotal = runningTotal + Totalone
'show the user the amount and price
picResults.Print Numberone; " Prawns:", , FormatCurrency(Totalone)
End Sub

Private Sub cmdMain2_Click()
'declare the variables
Dim Numbertwo As Integer, Totaltwo As Integer
'calculate how many Honey Meatballs the user wants
Numbertwo = InputBox("How many Honey Meatballs would you like?")
Totaltwo = 25 * Numbertwo
runningTotal = runningTotal + Totaltwo
'show the user the amount and price
picResults.Print Numbertwo; " Honey Meatballs:", FormatCurrency(Totaltwo)
End Sub

Private Sub cmdMain3_Click()
'declare the variables
Dim Numberthree As Integer, Totalthree As Integer
'calculate how many Kongbao Chickens the user wants
Numberthree = InputBox("How many Kongbao Chickens would you like?")
Totalthree = 18 * Numberthree
runningTotal = runningTotal + Totalthree
'show the user the amount and price
picResults.Print Numberthree; " Kongbao Chickens:", FormatCurrency(Totalthree)
End Sub

Private Sub cmdMain4_Click()
'declare the variables
Dim Numberfour As Integer, Totalfour As Integer
'calculate how many Lotus the user wants
Numberfour = InputBox("How many Tender Porks would you like?")
Totalfour = 20 * Numberfour
runningTotal = runningTotal + Totalfour
'show the user the amount and price
picResults.Print Numberfour; " Tender Porks:", FormatCurrency(Totalfour)
End Sub

Private Sub cmdNext_Click()
frmMain.Hide
frmDessert.Show
End Sub

Private Sub cmdOrder_Click()
'here user is able to order
'declare the variables
Dim Order As Integer, Ready As Integer, Tax As Single, Total As Single, Dishes(1 To 30) As String, Price(1 To 30) As Integer, CTR As Integer
'ask if the user is ready to order
Ready = InputBox("Input 1 if you are ready to order; otherwise if you want to look at the menu again")
If Ready = 1 Then
    'first open and sort the data
    Open App.Path & "\Dishes.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        'get a data set from the file
        Input #1, Dishes(CTR), Price(CTR)
    Loop

    MsgBox "Make sure you have memorized all the dishes' numbers!"
    Order = InputBox("Please enter the numbers of dishes you want(1 to 15);enter -999 to indicate end of ordering")
picResults.Print
picResults.Print "You would also like:"

    Do While Order <> -999
        picResults.Print Dishes(Order), FormatCurrency(Price(Order))
        runningTotal = runningTotal + Price(Order)
        Order = InputBox("Please enter the numbers of dishes you want;enter -999 to indicate end of ordering")
    Loop
Close #1

Tax = runningTotal * 0.08
Total = runningTotal + Tax

picResults.Print
picResults.Print "------------------------------------------------------------------------"
picResults.Print "Taxes:", FormatCurrency(Tax)
picResults.Print "Total:", FormatCurrency(Total)

Totalmaincost = Totalmaincost + Total
Totalcost = Totalcost + Totalmaincost
Else: MsgBox "Please check the menu again."
End If
        

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdRead_Click()
'declare the variables
Dim Order As Integer, OrderNow As Integer
'this opens the Dishes file in the folder, and helps user to order more.
Open App.Path & "\Dishes.txt" For Input As #1
'Initialize CTR with the value of zero
CTR = 0
'get ready for the menu
picResults.Cls
picResults.Print "We also have:"
'use a loop to show the user the dishes and ask the user to order
'from the menu.

Do While Not EOF(1)
        CTR = CTR + 1
        'get a data set from the file
        Input #1, Dishes(CTR), Price(CTR)
        
        'list all of the dishes and their price
        picResults.Print Dishes(CTR), FormatCurrency(Price(CTR)), CTR
              
    Loop
Close #1
End Sub

Private Sub cmdSelf_Click()
'ask the user to input a price under which he/she accepts
'declare the variables
Dim Ownprice As Integer

'we need to clear the picturebox first
picResults.Cls

'read the data from the file
Open App.Path & "\Dishes.txt" For Input As #1
'Initialize CTR with the value of zero
CTR = 0

'here user inputs the price
Ownprice = InputBox("Please input a price under which you are willing to pay")

'we divide the price into different cases
Select Case Ownprice
    Case Is <= 10
        MsgBox "Sorry, we don't have any dish under $10.", , "Oops"
    Case Is <= 15
        picResults.Print "Dishes under "; FormatCurrency(Ownprice); ":"
        Do While Not EOF(1)
        CTR = CTR + 1
        
        Input #1, Dishes(CTR), Price(CTR)
        
        'find the wanted dishes
        If Price(CTR) < Ownprice Then
        picResults.Print Dishes(CTR), FormatCurrency(Price(CTR)), CTR
        End If
        Loop
     Case Is <= 20
        picResults.Print "Dishes under "; FormatCurrency(Ownprice); ":"
        Do While Not EOF(1)
        CTR = CTR + 1
       
        Input #1, Dishes(CTR), Price(CTR)
        
        'find the wanted dishes
        If Price(CTR) < Ownprice Then
        picResults.Print Dishes(CTR), FormatCurrency(Price(CTR)), CTR
        End If
        
        Loop
     Case Is <= 25
        picResults.Print "Dishes under "; FormatCurrency(Ownprice); ":"
        Do While Not EOF(1)
        CTR = CTR + 1
       
        Input #1, Dishes(CTR), Price(CTR)
        
        'find the wanted dishes
        If Price(CTR) < Ownprice Then
        picResults.Print Dishes(CTR), FormatCurrency(Price(CTR)), CTR
        End If
        
        Loop
    Case Is <= 30
        picResults.Print "Dishes under "; FormatCurrency(Ownprice); ":"
        Do While Not EOF(1)
        CTR = CTR + 1
       
        Input #1, Dishes(CTR), Price(CTR)
        
        'find the wanted dishes
        If Price(CTR) < Ownprice Then
        picResults.Print Dishes(CTR), FormatCurrency(Price(CTR)), CTR
        End If
        
        Loop
    Case Else
        MsgBox "Cool, you can order anything we have!"
        Do While Not EOF(1)
        CTR = CTR + 1
        'get a data set from the file
        Input #1, Dishes(CTR), Price
        
        'list all of the dishes and their price
        picResults.Print Dishes(CTR), FormatCurrency(Price), CTR
              
        Loop
    End Select
Close #1
End Sub

Private Sub cmdSort_Click()
'declare the variables
Dim Pass As Integer, Pos As Integer, O As Integer, Tempprice As Integer

'compare the prices
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Price(Pos) > Price(Pos + 1) Then
        Tempprice = Price(Pos)
        Price(Pos) = Price(Pos + 1)
        Price(Pos + 1) = Tempprice
        End If
    Next Pos
Next Pass

'print out the dishes in order price
picResults.Cls
For O = 1 To CTR
    picResults.Print Dishes(O), FormatCurrency(Price(O), 2)
Next O
End Sub

Private Sub cmdUnder20_Click()

'declare the variables
Dim I As Integer
'clear the picturebox
picResults.Cls

'get ready for the menu under $20
picResults.Print "For the dishes under $20:"
'use a loop to show the user the dishes and ask the user to order
'from the menu.

For I = 1 To CTR
        
        'find the target dishes
        If Price(I) < 20 Then
        picResults.Print Dishes(I), FormatCurrency(Price(I)), I
        End If
         
Next I



End Sub

