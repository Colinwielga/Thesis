VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00400000&
   Caption         =   "CarProject"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   8775
      Left            =   2880
      ScaleHeight     =   8715
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   240
      Width           =   5895
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   8520
      Width           =   735
   End
   Begin VB.CommandButton cmdPic 
      Caption         =   "Show Picture"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "Display List of Affordable Cars"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Enter Financial Info"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input File"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      MaskColor       =   &H0000FFFF&
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer, Car(1 To 35) As String, Price(1 To 35) As Single
Dim Pass As Integer, Comp As Integer, TempCar As String, TempPrice As Single
Dim Wage As Single, Hours As Single, Expense As Single, SaveTime As Single, Monthly As Single, Budget As Single
Public ctr As Integer, Number As Integer
Public Path As String
'This program lets the user view vehicles that he can afford
'based on financial input from the user
Private Sub cmdClear_Click()
    'Clears the picture box
    picResults.Cls
    cmdInfo.Enabled = True
    cmdList.Enabled = False
    cmdPic.Enabled = False
End Sub

Private Sub cmdInfo_Click()
    'displays input box for the user to input his financial information
    cmdInput.Enabled = False
    ctr = 0
    picResults.Cls
    Wage = (InputBox("Enter your wage per hour:"))
    Hours = InputBox("Enter the number of hours you work in a typical week:")
    Expense = InputBox("Enter your expenses per month:")
    SaveTime = InputBox("Enter how many months you plan to save:")
    Monthly = ((Wage * Hours * 52) / 12) - Expense
    Budget = Monthly * SaveTime
    picResults.Print "After"; SaveTime; "months you will have saved "; FormatCurrency(Budget, 0); "."
    picResults.Print
    cmdList.Enabled = True
End Sub

Private Sub cmdInput_Click()
'Inputs the list of cars and their prices and sorts them according to ascending price
Path = "N:\CS130\handin\Gorres, Mark\"
    cmdInput.Enabled = False
    cmdInfo.Enabled = True
    cmdClear.Enabled = True
    'Open "M:\CS130\Project\carlist.txt" For Input As #1
    Open Path & "carlist.txt" For Input As #1
    For x = 1 To 35
        Input #1, Car(x), Price(x)
    Next x
    For Pass = 1 To 34
        For Comp = 1 To (35 - Pass)
            If Price(Comp) > Price(Comp + 1) Then
                TempPrice = Price(Comp)
                Price(Comp) = Price(Comp + 1)
                Price(Comp + 1) = TempPrice
                TempCar = Car(Comp)
                Car(Comp) = Car(Comp + 1)
                Car(Comp + 1) = TempCar
            End If
        Next Comp
    Next Pass
    Close #1
End Sub

Private Sub cmdList_Click()
'Calculates how much money the user will have to work with and
'displays a list of affordable cars
    cmdPic.Enabled = True
If Budget < 9400 Then
    picResults.Print "Perhaps you should look into a nice moped."
End If
If Budget >= 9400 Then
    picResults.Print "You will be able to afford the following vehicles:"
    picResults.Print
    picResults.Print "          Vehicle", Tab(45); "Price"
    picResults.Print "****************************************************************"
    For x = 1 To 35
    ctr = ctr + 1
    If Budget >= Price(x) Then
        picResults.Print ctr; ". "; Tab(7); Car(x), Tab(44); FormatCurrency(Price(x), 0)
    End If
    Next x
End If
cmdList.Enabled = False
End Sub

Private Sub cmdPic_Click()
'Displays a seperate form containing a picture of the car chosen
'by the viewer based off a number from an input box
    Number = InputBox("Enter the number of the vehicle you would like to see:")
    Select Case Number
        Case 1
            Kia.Show
        Case 2
            Aveo.Show
        Case 3
            Focus.Show
        Case 4
            Civic.Show
        Case 5
            Sentra.Show
        Case 6
            Frontier.Show
        Case 7
            Ranger.Show
        Case 8
            RSX.Show
        Case 9
            Jetta.Show
        Case 10
            Mazda6.Show
        Case 11
            Stratus.Show
        Case 12
            Accord.Show
        Case 13
            Altima.Show
        Case 14
            Escape.Show
        Case 15
            Liberty.Show
        Case 16
            Ram.Show
        Case 17
            F150.Show
        Case 18
            Chrysler300.Show
        Case 19
            Explorer.Show
        Case 20
            Nissan350Z.Show
        Case 21
            Trailblazer.Show
        Case 22
            A4.Show
        Case 23
            Expedition.Show
        Case 24
            Suburban.Show
        Case 25
            LS.Show
        Case 26
            Corvette.Show
        Case 27
            BMW.Show
        Case 28
            H2.Show
        Case 29
            Navigator.Show
        Case 30
            A8.Show
        Case 31
            Porsche.Show
        Case 32
            SL500.Show
        Case 33
            Gallardo.Show
        Case 34
            Maybach.Show
        Case 35
            Enzo.Show
        Case Is > 35
            MsgBox "Please input an appropriate number."
    End Select
    
End Sub

Private Sub cmdQuit_Click()
'Ends the program
    End
End Sub

