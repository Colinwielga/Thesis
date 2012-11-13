VERSION 5.00
Begin VB.Form PortfolioTracker 
   BackColor       =   &H80000005&
   Caption         =   "Portfolio Tracker - By Brandon Trampe"
   ClientHeight    =   10485
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureDisplayBox 
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   11955
      TabIndex        =   12
      Top             =   5040
      Width           =   12015
   End
   Begin VB.CommandButton cmdBL 
      Caption         =   "Today's ""Bottom Line"""
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   4560
      Width           =   5055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   10560
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Stock"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSummary 
      Caption         =   "Today's Portfolio Summary"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   4080
      Width           =   5055
   End
   Begin VB.CommandButton cmdSortG 
      Caption         =   "Sort By Today's Gain"
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   3600
      Width           =   3375
   End
   Begin VB.CommandButton cmdLosers 
      Caption         =   "Show Today's Losers"
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CommandButton cmdGainers 
      Caption         =   "Show Today's Gainers"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CommandButton cmdSortA 
      BackColor       =   &H80000007&
      Caption         =   "Sort Alphabetically"
      Height          =   375
      Left            =   3600
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   3600
      Width           =   5055
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Stocks"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   3375
   End
   Begin VB.PictureBox pbxResults 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   11955
      TabIndex        =   1
      Top             =   240
      Width           =   12015
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Portfolio Tracker v1.0"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "PortfolioTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PortfolioTracker v1.0 (portfoliotracker.vbp) '
' Portfolio Tracker - By Brandon Trampe '
' Written By Brandon Trampe '
' Written October 20-22, 2003 '
' This program and this form were written to help a person easily track their stock portfolio.
' Data is pulled from a flat text file (portfolio.txt) and displayed in a user-specified way.
' Stock data can be obtained and written to a text file using an open-source PHP content grabber.
' The grabber can be set to obtain data on whatever stocks the user specifies, and this program is set to handle any number of stocks up to 50.

Option Explicit
Dim company(1 To 50) As String
Dim ticker(1 To 50) As String
Dim price(1 To 50) As Double
Dim change(1 To 50) As Double
Dim percent(1 To 50) As Double
Dim shares(1 To 50) As Integer
Dim value(1 To 50) As Double
Dim changex(1 To 50) As Double
Dim strPath As String



Private Sub cmdBL_Click() ' Displays the Bottom Line, or the value and change in value of a user portfolio
    pbxResults.Cls
    Dim x As Integer
    Dim Q As Integer
    Dim absx As Double
    Dim totalchange As Double
    totalchange = 0
    Dim totalvalue As Double
    totalvalue = 0
    Dim strFile As String
    strFile = strPath & "portfolio.txt"
    Open strFile For Input As #1
    Do Until EOF(1)
        x = x + 1
        Input #1, company(x), ticker(x), shares(x), price(x), change(x), percent(x)
        value(x) = shares(x) * price(x)
        changex(x) = shares(x) * change(x)
        totalvalue = totalvalue + value(x)
        totalchange = totalchange + changex(x)
    Loop
    Close #1
    absx = Abs(totalchange)
    If totalchange > 0 Then
        pbxResults.Print "You gained "; FormatCurrency(totalchange, 2); " today and your portfolio is worth "; FormatCurrency(totalvalue, 2); "."
    Else
        pbxResults.Print "You lost "; FormatCurrency(absx, 2); " today and your portfolio is worth "; FormatCurrency(totalvalue, 2); "."
    End If
    
End Sub

Private Sub cmdClear_Click() ' Clears Display Box
    pbxResults.Cls
End Sub

Private Sub cmdFind_Click() ' Finds a user-specified stock and displays data for it
    pbxResults.Cls
    Dim x As Integer
    Dim findticker As String
    Dim totalvalue As Double
    totalvalue = 0
    Dim totalchange As Double
    totalchange = 0
    pbxResults.Print "Company", "Ticker", "Shares", "Price", "Change", "Percent", "Value", "Change"
    pbxResults.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim strFile As String
    strFile = strPath & "portfolio.txt"
    Open strFile For Input As #1
        Do Until EOF(1)
            x = x + 1
            Input #1, company(x), ticker(x), shares(x), price(x), change(x), percent(x)
            value(x) = shares(x) * price(x)
            changex(x) = shares(x) * change(x)
            totalvalue = totalvalue + value(x)
            totalchange = totalchange + changex(x)
        Loop
    findticker = InputBox("Enter A Ticker Symbol")
    x = 1
        Do Until ticker(x) = findticker Or x = 50
            x = x + 1
        Loop
    If x = 50 Then
        pbxResults.Print "This stock isn't in your portfolio."
    Else
        pbxResults.Print company(x), findticker, shares(x), FormatCurrency(price(x), 2), FormatCurrency(change(x), 2), percent(x); "%", FormatCurrency(value(x), 2), FormatCurrency(changex(x), 2)
    End If
   
   
   Close #1
End Sub

Private Sub cmdGainers_Click() ' Shows the stocks that appreciated in value on that day.
    pbxResults.Cls
    Dim y As Integer
    Dim x As Integer
    pbxResults.Print "Company", "Ticker", "Price", "Change", "Percent"
    pbxResults.Print "----------------------------------------------------------------------------------------------------------"
    Dim strFile As String
    strFile = strPath & "portfolio.txt"
    Open strFile For Input As #1
        Do Until EOF(1)
            x = x + 1
            Input #1, company(x), ticker(x), shares(x), price(x), change(x), percent(x)
                If change(x) > 0 Then
                    pbxResults.Print company(x), ticker(x), FormatCurrency(price(x), 2), FormatCurrency(change(x), 2), percent(x); "%"
                End If
        Loop
    Close #1
End Sub

Private Sub cmdLosers_Click() ' Shows stocks that depreciated in value that day.
    pbxResults.Cls
    Dim y As Integer
    Dim x As Integer
    pbxResults.Print "Company", "Ticker", "Price", "Change", "Percent"
    pbxResults.Print "----------------------------------------------------------------------------------------------------------"
    Dim strFile As String
    strFile = strPath & "portfolio.txt"
    Open strFile For Input As #1
    Do Until EOF(1)
        x = x + 1
        Input #1, company(x), ticker(x), shares(x), price(x), change(x), percent(x)
            If change(x) < 0 Then
                pbxResults.Print company(x), ticker(x), FormatCurrency(price(x), 2), FormatCurrency(change(x), 2), percent(x); "%"
            End If
        Loop
    Close #1
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdShow_Click() ' Shows all stocks in portfolio.
    pbxResults.Cls
    Dim x As Integer
    pbxResults.Print "Company", "Ticker", "Price", "Change", "Percent"
    pbxResults.Print "----------------------------------------------------------------------------------------------------------"
    Dim strFile As String
    strFile = strPath & "portfolio.txt"
    Open strFile For Input As #1
    Do Until EOF(1)
        x = x + 1
        Input #1, company(x), ticker(x), shares(x), price(x), change(x), percent(x)
        pbxResults.Print company(x), ticker(x), FormatCurrency(price(x), 2), FormatCurrency(change(x), 2), percent(x); "%"
    Loop
    Close #1
End Sub

Private Sub cmdSortA_Click() 'Sorts stocks alphabetically.
    Dim x As Integer
    Dim Pass As Integer
    Dim N As Integer
    Dim temp As String
    pbxResults.Cls
    Dim strFile As String
    strFile = strPath & "portfolio.txt"
    Open strFile For Input As #1
    Do Until EOF(1)
        x = x + 1
        Input #1, company(x), ticker(x), shares(x), price(x), change(x), percent(x)
    Loop
    pbxResults.Print "Company", "Ticker", "Price", "Change", "Percent"
    pbxResults.Print "----------------------------------------------------------------------------------------------------------"
    N = x
    For Pass = 1 To N - 1
        For x = 1 To N - Pass
            If company(x) > company(x + 1) Then
                temp = company(x)
                company(x) = company(x + 1)
                company(x + 1) = temp
                
                temp = ticker(x)
                ticker(x) = ticker(x + 1)
                ticker(x + 1) = temp
                
                temp = price(x)
                price(x) = price(x + 1)
                price(x + 1) = temp
                
                temp = change(x)
                change(x) = change(x + 1)
                change(x + 1) = temp
                
                temp = percent(x)
                percent(x) = percent(x + 1)
                percent(x + 1) = temp
            End If
        Next x
    Next Pass
    
    For x = 1 To N
        pbxResults.Print company(x), ticker(x), FormatCurrency(price(x), 2), FormatCurrency(change(x), 2), percent(x); "%"
    Next x
    
    Close #1
End Sub

Private Sub cmdSortG_Click() ' Sorts by gain in dollar value.
    Dim x As Integer
    Dim Pass As Integer
    Dim N As Integer
    Dim temp As String
    Dim strFile As String
    strFile = strPath & "portfolio.txt"
    Open strFile For Input As #1
    Do Until EOF(1)
        x = x + 1
        Input #1, company(x), ticker(x), shares(x), price(x), change(x), percent(x)
    Loop
    Close #1
    pbxResults.Cls
    pbxResults.Print "Company", "Ticker", "Price", "Change", "Percent"
    pbxResults.Print "----------------------------------------------------------------------------------------------------------"
    N = x
    For Pass = 1 To N - 1
        For x = 1 To N - Pass
            If change(x) < change(x + 1) Then
                temp = change(x)
                change(x) = change(x + 1)
                change(x + 1) = temp
                
                temp = ticker(x)
                ticker(x) = ticker(x + 1)
                ticker(x + 1) = temp
                
                temp = price(x)
                price(x) = price(x + 1)
                price(x + 1) = temp
                
                temp = company(x)
                company(x) = company(x + 1)
                company(x + 1) = temp
                
                temp = percent(x)
                percent(x) = percent(x + 1)
                percent(x + 1) = temp
            End If
        Next x
    Next Pass
    
    For x = 1 To N
        pbxResults.Print company(x), ticker(x), FormatCurrency(price(x), 2), FormatCurrency(change(x), 2), percent(x); "%"
    Next x
End Sub

Private Sub cmdSummary_Click() ' Shows a detailed summary of the portfolio.
    pbxResults.Cls
    Dim x As Integer
    Dim totalvalue As Double
    totalvalue = 0
    Dim totalchange As Double
    totalchange = 0
    pbxResults.Print "Company", "Ticker", "Shares", "Price", "Change", "Percent", "Value", "Change"
    pbxResults.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim strFile As String
    strFile = strPath & "portfolio.txt"
    Dim strPicture As String
    strPicture = strPath & "money.jpg"
    Open strFile For Input As #1
        Do Until EOF(1)
            x = x + 1
            Input #1, company(x), ticker(x), shares(x), price(x), change(x), percent(x)
            value(x) = shares(x) * price(x)
            changex(x) = shares(x) * change(x)
            totalvalue = totalvalue + value(x)
            totalchange = totalchange + changex(x)
            pbxResults.Print company(x), ticker(x), shares(x), FormatCurrency(price(x), 2), FormatCurrency(change(x), 2), percent(x); "%", FormatCurrency(value(x), 2), FormatCurrency(changex(x), 2)
        Loop
        pbxResults.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        pbxResults.Print Tab(72); "TOTALS", FormatCurrency(totalvalue, 2), FormatCurrency(totalchange, 2)
    Close #1
    If totalchange > 0 Then
        MsgBox "Congratulations, you made money today!", , "Congrats!"
        PictureDisplayBox.Picture = LoadPicture(strPicture)
    Else
        MsgBox "Sorry, you lost money today.  There is always tomorrow!", , "Oh No!"
    End If
    
    
    
End Sub



   
Private Sub Form_Load()
    strPath = "N:\CS130\handin\bdtrampe\"
End Sub
