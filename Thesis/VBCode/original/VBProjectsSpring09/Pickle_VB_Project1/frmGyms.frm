VERSION 5.00
Begin VB.Form frmGyms 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   10695
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H0080FF80&
      Caption         =   "Previous Screen"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton cmdEstimate 
      BackColor       =   &H008080FF&
      Caption         =   "Estimate Cost"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9240
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   9000
      Picture         =   "frmGyms.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   3075
      TabIndex        =   19
      Top             =   9120
      Width           =   3135
   End
   Begin VB.PictureBox picGymNameThree 
      Height          =   1815
      Left            =   5400
      Picture         =   "frmGyms.frx":0A86
      ScaleHeight     =   1755
      ScaleWidth      =   4155
      TabIndex        =   17
      Top             =   6600
      Width           =   4215
   End
   Begin VB.PictureBox picGymNameTwo 
      BackColor       =   &H00C0C0FF&
      Height          =   2295
      Left            =   10560
      Picture         =   "frmGyms.frx":1DC5
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   18
      Top             =   6600
      Width           =   2295
   End
   Begin VB.PictureBox picGymNameOne 
      BackColor       =   &H00800080&
      Height          =   1575
      Left            =   6120
      Picture         =   "frmGyms.frx":262B
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   16
      Top             =   8760
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddAGym 
      BackColor       =   &H008080FF&
      Caption         =   "Add A Gym"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   4095
   End
   Begin VB.CommandButton cmdTotalSoFar 
      BackColor       =   &H0080FF80&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0080FF80&
      Caption         =   "Take me to my final total"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox txtGyms 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   1080
      TabIndex        =   12
      Text            =   "Add A Gym Membership?!"
      Top             =   120
      Width           =   12975
   End
   Begin VB.PictureBox picResultsTwo 
      BackColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   5400
      ScaleHeight     =   1155
      ScaleWidth      =   9195
      TabIndex        =   11
      Top             =   5280
      Width           =   9255
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H008080FF&
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9240
      Width           =   2055
   End
   Begin VB.CommandButton cmdReadFile 
      BackColor       =   &H008080FF&
      Caption         =   "Read the File"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox txtNumberofMonths 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Left            =   600
      TabIndex        =   6
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox txtSelectGym 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Left            =   600
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FF80&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortLowtoHigh 
      BackColor       =   &H008080FF&
      Caption         =   "Sort Prices (Low to High)"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   4095
   End
   Begin VB.CommandButton cmdSearchPrice 
      BackColor       =   &H008080FF&
      Caption         =   "Search by Price"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   4095
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H008080FF&
      Caption         =   "Show Gyms"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   4095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      Height          =   3735
      Left            =   5400
      ScaleHeight     =   3675
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   1320
      Width           =   9255
   End
   Begin VB.Label lblMonths 
      BackColor       =   &H0080C0FF&
      Caption         =   "How many months would you like the membership?"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   8
      Top             =   7800
      Width           =   3135
   End
   Begin VB.Label lblGym 
      BackColor       =   &H0080C0FF&
      Caption         =   "What number gym would you like a membership to?"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   7
      Top             =   6240
      Width           =   3135
   End
End
Attribute VB_Name = "frmGyms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name of Project: Build Your Own Home Gym
'Form Name: frmGyms
'Author: Michelle Pickle
'Written on: March 12th 2009
'The objective of this form is to show the user the various cost of belonging to gyms.  The user can choose to purchase a gym membership in addition to their own equipment.
    'The user is able to search and sort through these gyms in a variety of ways.

Option Explicit
'the varaibles declared here are those used in multiple subroutines(multiple buttons)
Dim Number(1 To 50) As Integer
Dim NameofGym(1 To 50) As String
Dim PricePerMonth(1 To 50) As Double
Dim PriceCount As Integer

Private Sub cmdAddAGym_Click()
'this button is meant to append a file aka or add a gym to the end of the file
Dim Name As String
Dim NextNumber As Integer
Dim Price As Double

    Open App.Path & "\gymprices.txt" For Append As #1
        NextNumber = InputBox("Please Enter the Next Number In Sequence (i.e. 11,12,13,etc...)", "Enter Number")
        Name = InputBox("Please Enter A Gym Name", "Gym Name")
        Price = InputBox("Please Enter Price Per Month", "Price")
    
    Write #1, NextNumber, Name, Price
    Close
    
'when this action is complete, these buttons will or will not be available to use
    cmdDisplay.Visible = True
    cmdReadFile.Visible = True
    cmdSearchPrice.Visible = False
    cmdSortLowtoHigh.Visible = False
       
    
End Sub

Private Sub cmdCalculate_Click()
'this button allows the user to calulate the information they entered in the text boxes (what gym they want to belong to, and for how many months
'the variables used in this subroutine are declared
    Dim GymType As Integer
    Dim Months As Integer
    Dim Subtotal As Double
'the number the user eneters in the text box are assigned to certain variables
    GymType = txtSelectGym.Text
    Months = txtNumberofMonths.Text
    picResultsTwo.Cls
'using the select case exmaples, lets us declare what will happen depending on what number the user enters in the first textbok
    Select Case GymType
        'if the number is a one
        Case Is = 1
            'the price of the gym per month times the number of months the user entered he/she wanted to belong is declared as the subtotal
            gymtotal = 54.9 * Months
            'prints a sentence decribing what the user entered and the total priced formatted in money
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            'add the users selection to the running total
            runningtotal = runningtotal + gymtotal
        Case Is = 2
            gymtotal = 34.95 * Months
            picResultsTwo.Print "PUCHARSE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
        Case Is = 3
            gymtotal = 49.5 * Months
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
        Case Is = 4
            gymtotal = 37.95 * Months
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
        Case Is = 5
            gymtotal = 53.99 * Months
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
        Case Is = 6
            gymtotal = 51.95 * Months
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
        Case Is = 7
            gymtotal = 34.5 * Months
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
        Case Is = 8
            gymtotal = 29.99 * Months
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
        Case Is = 9
            gymtotal = 32.5 * Months
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
        Case Is = 10
            gymtotal = 72.87 * Months
            picResultsTwo.Print "PURCHASE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
            runningtotal = runningtotal + gymtotal
         Case Else
            MsgBox "The number you entered is invalid.  Please enter another number", , "Error"
    End Select
     
End Sub

Private Sub cmdEstimate_Click()
    'this button allows the user to calulate the information they entered in the text boxes (what gym they want to belong to, and for how many months
'the variables used in this subroutine are declared
    Dim GymType As Integer
    Dim Months As Integer
    Dim Subtotal As Double
'the number the user eneters in the text box are assigned to certain variables
    GymType = txtSelectGym.Text
    Months = txtNumberofMonths.Text
    picResultsTwo.Cls
'using the select case exmaples, lets us declare what will happen depending on what number the user enters in the first textbok
    Select Case GymType
        'if the number is a one
        Case Is = 1
            'the price of the gym per month times the number of months the user entered he/she wanted to belong is declared as the subtotal
            gymtotal = 54.9 * Months
            'prints a sentence decribing what the user entered and the total priced formatted in money
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 2
            gymtotal = 34.95 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 3
            gymtotal = 49.5 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 4
            gymtotal = 37.95 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 5
            gymtotal = 53.99 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 6
            gymtotal = 51.95 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 7
            gymtotal = 34.5 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 8
            gymtotal = 29.99 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 9
            gymtotal = 32.5 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
        Case Is = 10
            gymtotal = 72.87 * Months
            picResultsTwo.Print "ESTIMATE: Your total for"; Months; "months is "; FormatCurrency(gymtotal)
         Case Else
            MsgBox "The number you entered is invalid.  Please enter another number", , "Error"
    End Select
     
End Sub

Private Sub cmdContinue_Click()
'changes forms
    frmGyms.Hide
    frmReceipt.Show
End Sub

Private Sub cmdDisplay_Click()
'declares the variables
    Dim K As Double
'when display is clicked, the user is no longer able to read the file.  The user does not need to read the file, because it has already been read
    cmdReadFile.Visible = False
'clears the results in the picture box
    picResults.Cls
'prints the headings, with a line of stars underneath to allow the user to read the data more clearly
    picResults.Print "Number", "Gym"; Tab(40); "Price Per Month"
    picResults.Print "*********************************************************************"
'the for-next loop goes through the read array and prints the number (1,2,3...) associated with each gym as well as the name of the gym, and the price per month.
    For K = 1 To PriceCount
        picResults.Print Number(K), NameofGym(K); Tab(40); FormatCurrency(PricePerMonth(K))
    Next K
'once the button is clicked, then the remaining two command buttons become visible
cmdSortLowtoHigh.Visible = True
cmdSearchPrice.Visible = True
cmdAddAGym.Visible = True
txtSelectGym.Visible = True
txtNumberofMonths.Visible = True
cmdEstimate.Visible = True
cmdCalculate.Visible = True
End Sub

Private Sub cmdPrevious_Click()
'allows the user to navigate to the previous screen
    frmMachines.Visible = True
    frmGyms.Visible = False
End Sub

Private Sub cmdQuit_Click()
'this ends the program
    End
End Sub

Private Sub cmdReadFile_Click()
'this button only reads the file, no data is printed
'only the display button becomes visible after the ReadFile button is selected
    cmdDisplay.Visible = True
    cmdSortLowtoHigh.Visible = False
    cmdSearchPrice.Visible = False
    cmdAddAGym.Visible = False
    txtSelectGym.Visible = False
    txtNumberofMonths.Visible = False
    cmdEstimate.Visible = False
    cmdCalculate.Visible = False
    
'opens the file and prepars it to be read
   Open App.Path & "\gymprices.txt" For Input As #1
   PriceCount = 0
        Do Until EOF(1)
            PriceCount = PriceCount + 1
'this assigns a name to each column of data
            Input #1, Number(PriceCount), NameofGym(PriceCount), PricePerMonth(PriceCount)
        Loop
'Once the reading of the file is done, the pop up message box alerts the reader to the fact the reading of the file has benn completed
     MsgBox "Your file has been read", , "Read"
'the file is then closed
    Close #1
End Sub

Private Sub cmdSearchPrice_Click()
'variables are declared
    Dim SPrice As Double
    Dim K As Double
'the buttons ReadFile and Display are not needed when searching for a price, so they become not visible
    cmdReadFile.Visible = False
    cmdDisplay.Visible = False
'clears the Result box
    picResults.Cls
'the user enter the max. price he/she is willing to spend per month on belonging to a fitness center.  This is then set equal to "SPrice"
    SPrice = InputBox("Please enter the highest amount willing to spend per month on a gym membership", "Enter Amount", "Enter Amount")
'the headings are printed in the result box
    picResults.Print "Number", "Name of Gym"; Tab(40); "Price Per Month"
'star line printed making the data easy to read, and in a clean and neat order
    picResults.Print "*****************************************************"
'this for-next loop cycles through all of the information in the array.  The If-Then loop looks for all the prices lower the the price entered by the user
    For K = 1 To PriceCount
        If PricePerMonth(K) < SPrice Then
'this prints only the gym and the price of belonging to the gym that are under the price entered by the user.
            picResults.Print Number(K), NameofGym(K); Tab(40); FormatCurrency(PricePerMonth(K))
        End If
    Next K
        
End Sub

Private Sub cmdSortLowtoHigh_Click()
'this buttons sorts the prices from low to high
'variables for this sub routine are declared
Dim tempPricePerMonth As Double
Dim tempNumber As Double
Dim tempNameofGym As String
Dim Pass As Double
Dim K As Double
Dim I As Double
'the picture box is cleared of any data previously present
picResults.Cls
'the buttons "ReadFile" is not visible, but the button Display is.  This user must click on the display button in order to view the sorted list
cmdReadFile.Visible = False
cmdDisplay.Visible = False
cmdSearchPrice.Visible = True
'keeps track of how many passes are made
    For Pass = 1 To PriceCount - 1
'keeps track of how many comparisons between numbers is being made
        For K = 1 To PriceCount - Pass
'declares condition of how the sorting is to be done (from low to high)
            If PricePerMonth(K) > PricePerMonth(K + 1) Then
'the remaining steps swithces the vaule if it is out of order
'this must be done will all variable because if one switches, all columns must switch or the data will not longer be accurate.
                tempPricePerMonth = PricePerMonth(K)
                PricePerMonth(K) = PricePerMonth(K + 1)
                PricePerMonth(K + 1) = tempPricePerMonth
                tempNumber = Number(K)
                Number(K) = Number(K + 1)
                Number(K + 1) = tempNumber
                tempNameofGym = NameofGym(K)
                NameofGym(K) = NameofGym(K + 1)
                NameofGym(K + 1) = tempNameofGym
            End If
        Next K
    Next Pass
'prints the heading
picResults.Print "Number", "Gym"; Tab(40); "Price Per Month"
picResults.Print "*******************************************************"
'prints the sorted data
For I = 1 To PriceCount
    picResults.Print Number(I), NameofGym(I); Tab(40); FormatCurrency(PricePerMonth(I))
Next I
End Sub

Private Sub cmdTotal_Click()
'switches forms
    frmGyms.Hide
    frmReceipt.Show
End Sub

Private Sub cmdTotalSoFar_Click()
'clears the picture box so no other data is present
    picResultsTwo.Cls
'allows the user to view their overall total of all the items they have purchased so far
    picResultsTwo.Cls
    picResultsTwo.Print "Your current total for all screens is "; FormatCurrency(runningtotal)
End Sub


Private Sub Form_Load()
'This code centers the form on computer screen upon loading.
'this code discovered from Cassie Scherer and Jordan Schmaltz project of developing a vacation

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

'When the form first appears, the only button that will be visible is the read file.  This is because in order to perform any of the remaining buttons the file must first be read
    cmdReadFile.Visible = True
    cmdDisplay.Visible = False
    cmdSearchPrice.Visible = False
    cmdSortLowtoHigh.Visible = False
    cmdAddAGym.Visible = False
    txtSelectGym.Visible = False
    txtNumberofMonths.Visible = False
    cmdEstimate.Visible = False
    cmdCalculate.Visible = False
End Sub


