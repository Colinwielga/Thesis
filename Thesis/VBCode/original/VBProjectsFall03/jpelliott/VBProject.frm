VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "Equity Sorter 1.0"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   FillColor       =   &H000000FF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   240
      Picture         =   "VBPROJ~1.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   3795
      TabIndex        =   17
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   1095
      Left            =   360
      MaskColor       =   &H00000000&
      TabIndex        =   13
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Portfolio Return"
      Height          =   1095
      Left            =   7560
      TabIndex        =   12
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   1095
      Left            =   2760
      TabIndex        =   11
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   5160
      TabIndex        =   10
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New Security To Memory"
      Height          =   1095
      Left            =   5160
      TabIndex        =   9
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Current Holdings"
      Height          =   1095
      Left            =   2760
      TabIndex        =   8
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortGLP 
      Caption         =   "% Gain/Loss"
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGL 
      Caption         =   "Gain/Loss"
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSortValue 
      Caption         =   "Current Value"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSortCost 
      Caption         =   "Acquisition Cost"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSortShares 
      Caption         =   "Shares"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSortTicker 
      Caption         =   "Ticker"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdSortName 
      Caption         =   "Company"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.PictureBox pbxResults 
      Height          =   3975
      Left            =   240
      ScaleHeight     =   3915
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   1920
      Width           =   9615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   "Click on Column Headings to Sort the Equities"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Please First Click LOAD to Load Portfolio"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   14
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Company(1 To 50) As String
Dim Ticker(1 To 50) As String
Dim Shares(1 To 50) As Integer
Dim AcquisitionCost(1 To 50) As Single
Dim CurrentValue(1 To 50) As Single
Dim GainLoss(1 To 50) As Single
Dim PercentGainLoss(1 To 50) As Single
Dim icount As Integer                               'Problem with negative nubmer colors (tried under sort name)
Dim strPath As String

Private Sub cmdAdd_Click()              'ADD IN MORE STOCKS TO PORTFOLIO
Dim C As Integer
'Dim strFile As String
'strFile = strPath & "data.txt"
'strPath = "N:\CS130\handin\jpelliott\"
Dim NumSecurity As Integer

'Open strFile For Output As #2
NumSecurity = InputBox("Please enter the number of securities you wish to add")
    For C = 1 To NumSecurity
        icount = icount + 1
        'InputBox ("Please enter A Security Name,Ticker Symbol,Number of Shares, Acquisition Cost, Current Value, Gain/Loss, and Percent Gain/Loss all seperated with commas")
        Company(icount) = InputBox("Please Enter A Security Name Beginning With An Uppercase Letter", "Name")
        Ticker(icount) = InputBox("Please Enter the Ticker Symbol all in UPPER CASE", "Ticker")
        Shares(icount) = InputBox("Please Enter the Number of Shares", "Shares")
        AcquisitionCost(icount) = InputBox("Please Enter the Acquisition Cost", "Acq.Cost")
        CurrentValue(icount) = InputBox("Please Enter the Total Current Value of the Securities", "Total Current Value")
        GainLoss(icount) = InputBox("Please Enter the Unrealized Gain/Loss on the Security", "Unrealized G/L")
        PercentGainLoss(icount) = InputBox("Please Enter the Unrealized Gain/Loss Percentage", "Unrealized G/L Percentage")
    Next C
Close #2
End Sub
Private Sub cmdClear_Click()
    MsgBox "You are clearing the Picture Box", , "Clear"
    pbxResults.Cls
End Sub

Private Sub cmdLoad_Click()
Dim C As Integer
icount = 0
C = 1
                            '***********HAVING PROBLEMS WITH ALL OF THIS NEW FILE PATH,*******
                            '***********SUGGESTED I LEAVE IT AS IS TO NOT MESS UP PROGRAM*****
'strPath = "N:\CS130\handin\jpelliott\"
'Dim strFile As String
'strFile = strPath & "data.txt"
'Open strFile For Input As #1

Open "N:\CS130\handin\jpelliott\data.txt" For Input As #1

Do Until EOF(1)
        Input #1, Company(C), Ticker(C), Shares(C), AcquisitionCost(C), CurrentValue(C), GainLoss(C), PercentGainLoss(C)
        icount = icount + 1
        C = C + 1
Loop
Close #1
End Sub

Private Sub cmdGL_Click()
Dim C As Integer
C = 1
Dim Pass As Integer
Dim temp As String
    
For Pass = 1 To icount - 1
    For C = 1 To icount - Pass
        If GainLoss(C) > GainLoss(C + 1) Then
            temp = Company(C)
            Company(C) = Company(C + 1)
            Company(C + 1) = temp
            temp = Ticker(C)
            Ticker(C) = Ticker(C + 1)
            Ticker(C + 1) = temp
            temp = Shares(C)
            Shares(C) = Shares(C + 1)
            Shares(C + 1) = temp
            temp = AcquisitionCost(C)
            AcquisitionCost(C) = AcquisitionCost(C + 1)
            AcquisitionCost(C + 1) = temp
            temp = CurrentValue(C)
            CurrentValue(C) = CurrentValue(C + 1)
            CurrentValue(C + 1) = temp
            temp = GainLoss(C)
            GainLoss(C) = GainLoss(C + 1)
            GainLoss(C + 1) = temp
            temp = PercentGainLoss(C)
            PercentGainLoss(C) = PercentGainLoss(C + 1)
            PercentGainLoss(C + 1) = temp
          End If
    Next C
Next Pass

pbxResults.Print " "
For C = 1 To icount
    pbxResults.Print Company(C), Tab(40); Ticker(C); Tab(60); Shares(C), FormatCurrency(AcquisitionCost(C), 2), FormatCurrency(CurrentValue(C), 2), FormatCurrency(GainLoss(C), 2), FormatPercent(PercentGainLoss(C), 2)
Next C
End Sub

Private Sub cmdDisplay_Click()
Dim C As Integer
C = 1

For C = 1 To icount
    pbxResults.Print Company(C), Tab(40); Ticker(C); Tab(60); Shares(C), FormatCurrency(AcquisitionCost(C), 2), FormatCurrency(CurrentValue(C), 2), FormatCurrency(GainLoss(C), 2), FormatPercent(PercentGainLoss(C), 2)
Next C
End Sub

Private Sub cmdQuit_Click()
    Form1.Hide
    Form2.Show
End Sub


Private Sub cmdReturn_Click()           'CALCULATE THE UNREALIZED RETURN
                                        
'1. sumGL = gain/loss (Add them all up)
'2. sumAC = Acquisition Cost (Add up what they cost)
'3. sumGL/SumAC (Gives me my % GL) for portfolio

Dim C As Integer
Dim L As Integer
Dim sumGL As Single
sumGL = 0
Dim sumAC As Single

For C = 1 To icount
    sumGL = sumGL + GainLoss(C)
Next C

For C = 1 To icount
    sumAC = sumAC + AcquisitionCost(C)
Next C

pbxResults.Print " "
pbxResults.Print "Unrealized Gain/Loss", FormatPercent(sumGL / sumAC)


End Sub

Private Sub cmdSortCost_Click()
Dim C As Integer
C = 1
Dim Pass As Integer
Dim temp As String
    
For Pass = 1 To icount - 1
    For C = 1 To icount - Pass
        If AcquisitionCost(C) > AcquisitionCost(C + 1) Then
            temp = Company(C)
            Company(C) = Company(C + 1)
            Company(C + 1) = temp
            temp = Ticker(C)
            Ticker(C) = Ticker(C + 1)
            Ticker(C + 1) = temp
            temp = Shares(C)
            Shares(C) = Shares(C + 1)
            Shares(C + 1) = temp
            temp = AcquisitionCost(C)
            AcquisitionCost(C) = AcquisitionCost(C + 1)
            AcquisitionCost(C + 1) = temp
            temp = CurrentValue(C)
            CurrentValue(C) = CurrentValue(C + 1)
            CurrentValue(C + 1) = temp
            temp = GainLoss(C)
            GainLoss(C) = GainLoss(C + 1)
            GainLoss(C + 1) = temp
            temp = PercentGainLoss(C)
            PercentGainLoss(C) = PercentGainLoss(C + 1)
            PercentGainLoss(C + 1) = temp
          End If
    Next C
Next Pass

pbxResults.Print " "
For C = 1 To icount
    pbxResults.Print Company(C), Tab(40); Ticker(C); Tab(60); Shares(C), FormatCurrency(AcquisitionCost(C), 2), FormatCurrency(CurrentValue(C), 2), FormatCurrency(GainLoss(C), 2), FormatPercent(PercentGainLoss(C), 2)
Next C
End Sub

Private Sub cmdSortGLP_Click()
Dim C As Integer
C = 1
Dim Pass As Integer
Dim temp As String
    
For Pass = 1 To icount - 1
    For C = 1 To icount - Pass
        If PercentGainLoss(C) > PercentGainLoss(C + 1) Then
            temp = Company(C)
            Company(C) = Company(C + 1)
            Company(C + 1) = temp
            temp = Ticker(C)
            Ticker(C) = Ticker(C + 1)
            Ticker(C + 1) = temp
            temp = Shares(C)
            Shares(C) = Shares(C + 1)
            Shares(C + 1) = temp
            temp = AcquisitionCost(C)
            AcquisitionCost(C) = AcquisitionCost(C + 1)
            AcquisitionCost(C + 1) = temp
            temp = CurrentValue(C)
            CurrentValue(C) = CurrentValue(C + 1)
            CurrentValue(C + 1) = temp
            temp = GainLoss(C)
            GainLoss(C) = GainLoss(C + 1)
            GainLoss(C + 1) = temp
            temp = PercentGainLoss(C)
            PercentGainLoss(C) = PercentGainLoss(C + 1)
            PercentGainLoss(C + 1) = temp
          End If
    Next C
Next Pass

pbxResults.Print " "
For C = 1 To icount
   pbxResults.Print Company(C), Tab(40); Ticker(C); Tab(60); Shares(C), FormatCurrency(AcquisitionCost(C), 2), FormatCurrency(CurrentValue(C), 2), FormatCurrency(GainLoss(C), 2), FormatPercent(PercentGainLoss(C), 2)
Next C
End Sub

Private Sub cmdSortName_Click()
Dim C As Integer
C = 1
Dim Pass As Integer
Dim temp As String
Dim forColor As String

    
For Pass = 1 To icount - 1
    For C = 1 To icount - Pass
        If Company(C) > Company(C + 1) Then
            temp = Company(C)
            Company(C) = Company(C + 1)
            Company(C + 1) = temp
            temp = Ticker(C)
            Ticker(C) = Ticker(C + 1)
            Ticker(C + 1) = temp
            temp = Shares(C)
            Shares(C) = Shares(C + 1)
            Shares(C + 1) = temp
            temp = AcquisitionCost(C)
            AcquisitionCost(C) = AcquisitionCost(C + 1)
            AcquisitionCost(C + 1) = temp
            temp = CurrentValue(C)
            CurrentValue(C) = CurrentValue(C + 1)
            CurrentValue(C + 1) = temp
            temp = GainLoss(C)
            GainLoss(C) = GainLoss(C + 1)
            GainLoss(C + 1) = temp
            temp = PercentGainLoss(C)
            PercentGainLoss(C) = PercentGainLoss(C + 1)
            PercentGainLoss(C + 1) = temp
          End If
    Next C
Next Pass
                                                'I CAN'T GET THIS PART TO WORK!!! TO CHANGE COLORS!!!CASEY SAID JUST ABOUT IMPOSSIBLE @ THIS LEVEL TO DO
                                                
pbxResults.Print " "
For C = 1 To icount
    pbxResults.Print Company(C), Tab(40); Ticker(C); Tab(60); Shares(C), FormatCurrency(AcquisitionCost(C), 2), FormatCurrency(CurrentValue(C), 2), ;
        'If GainLoss(C) < 1 Then
        '    forColor = &HFF&
            pbxResults.Print FormatCurrency(GainLoss(C), 2), ; 'RGB(Rnd * 256, Rnd * 0, Rnd * 0); FormatCurrency(GainLoss(C), 2), ;
        'Else
        '    forColor = &H0&
        '    pbxResults.Print FormatCurrency(GainLoss(C), 2), ; 'ForeColor = vbBlack; , 'RGB(Rnd * 256, Rnd * 256, Rnd * 256);
        'End If
    pbxResults.Print FormatPercent(PercentGainLoss(C), 2)
Next C
        
End Sub

Private Sub cmdSortShares_Click()
Dim C As Integer
C = 1
Dim Pass As Integer
Dim temp As String
    
For Pass = 1 To icount - 1
    For C = 1 To icount - Pass
        If Shares(C) > Shares(C + 1) Then
            temp = Company(C)
            Company(C) = Company(C + 1)
            Company(C + 1) = temp
            temp = Ticker(C)
            Ticker(C) = Ticker(C + 1)
            Ticker(C + 1) = temp
            temp = Shares(C)
            Shares(C) = Shares(C + 1)
            Shares(C + 1) = temp
            temp = AcquisitionCost(C)
            AcquisitionCost(C) = AcquisitionCost(C + 1)
            AcquisitionCost(C + 1) = temp
            temp = CurrentValue(C)
            CurrentValue(C) = CurrentValue(C + 1)
            CurrentValue(C + 1) = temp
            temp = GainLoss(C)
            GainLoss(C) = GainLoss(C + 1)
            GainLoss(C + 1) = temp
            temp = PercentGainLoss(C)
            PercentGainLoss(C) = PercentGainLoss(C + 1)
            PercentGainLoss(C + 1) = temp
          End If
    Next C
Next Pass

pbxResults.Print " "
For C = 1 To icount
    pbxResults.Print Company(C), Tab(40); Ticker(C); Tab(60); Shares(C), FormatCurrency(AcquisitionCost(C), 2), FormatCurrency(CurrentValue(C), 2), FormatCurrency(GainLoss(C), 2), FormatPercent(PercentGainLoss(C), 2)
Next C
End Sub

Private Sub cmdSortTicker_Click()
Dim C As Integer
C = 1
Dim Pass As Integer
Dim temp As String
    
For Pass = 1 To icount - 1
    For C = 1 To icount - Pass
        If Ticker(C) > Ticker(C + 1) Then
            temp = Company(C)
            Company(C) = Company(C + 1)
            Company(C + 1) = temp
            temp = Ticker(C)
            Ticker(C) = Ticker(C + 1)
            Ticker(C + 1) = temp
            temp = Shares(C)
            Shares(C) = Shares(C + 1)
            Shares(C + 1) = temp
            temp = AcquisitionCost(C)
            AcquisitionCost(C) = AcquisitionCost(C + 1)
            AcquisitionCost(C + 1) = temp
            temp = CurrentValue(C)
            CurrentValue(C) = CurrentValue(C + 1)
            CurrentValue(C + 1) = temp
            temp = GainLoss(C)
            GainLoss(C) = GainLoss(C + 1)
            GainLoss(C + 1) = temp
            temp = PercentGainLoss(C)
            PercentGainLoss(C) = PercentGainLoss(C + 1)
            PercentGainLoss(C + 1) = temp
          End If
    Next C
Next Pass

pbxResults.Print " "
For C = 1 To icount
    pbxResults.Print Company(C), Tab(40); Ticker(C); Tab(60); Shares(C), FormatCurrency(AcquisitionCost(C), 2), FormatCurrency(CurrentValue(C), 2), FormatCurrency(GainLoss(C), 2), FormatPercent(PercentGainLoss(C), 2)
Next C

End Sub

Private Sub cmdSortValue_Click()
Dim C As Integer
C = 1
Dim Pass As Integer
Dim temp As String
    
For Pass = 1 To icount - 1
    For C = 1 To icount - Pass
        If CurrentValue(C) > CurrentValue(C + 1) Then
            temp = Company(C)
            Company(C) = Company(C + 1)
            Company(C + 1) = temp
            temp = Ticker(C)
            Ticker(C) = Ticker(C + 1)
            Ticker(C + 1) = temp
            temp = Shares(C)
            Shares(C) = Shares(C + 1)
            Shares(C + 1) = temp
            temp = AcquisitionCost(C)
            AcquisitionCost(C) = AcquisitionCost(C + 1)
            AcquisitionCost(C + 1) = temp
            temp = CurrentValue(C)
            CurrentValue(C) = CurrentValue(C + 1)
            CurrentValue(C + 1) = temp
            temp = GainLoss(C)
            GainLoss(C) = GainLoss(C + 1)
            GainLoss(C + 1) = temp
            temp = PercentGainLoss(C)
            PercentGainLoss(C) = PercentGainLoss(C + 1)
            PercentGainLoss(C + 1) = temp
          End If
    Next C
Next Pass

pbxResults.Print " "
For C = 1 To icount
    pbxResults.Print Company(C), Tab(40); Ticker(C); Tab(60); Shares(C), FormatCurrency(AcquisitionCost(C), 2), FormatCurrency(CurrentValue(C), 2), FormatCurrency(GainLoss(C), 2), FormatPercent(PercentGainLoss(C), 2)
Next C
End Sub

