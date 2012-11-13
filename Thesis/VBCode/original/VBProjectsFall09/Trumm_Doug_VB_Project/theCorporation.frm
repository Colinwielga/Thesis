VERSION 5.00
Begin VB.Form frmWorkStation 
   BackColor       =   &H00000040&
   Caption         =   "Keeping track of finances"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDilbert 
      Height          =   4935
      Left            =   7560
      Picture         =   "theCorporation.frx":0000
      ScaleHeight     =   3087.886
      ScaleMode       =   0  'User
      ScaleWidth      =   5955
      TabIndex        =   8
      Top             =   2400
      Width           =   6015
   End
   Begin VB.CommandButton cmdPromotion 
      Caption         =   "Accept Promotion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11640
      TabIndex        =   7
      Top             =   8280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdSwitchBack 
      Caption         =   "Return to Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9600
      TabIndex        =   6
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7800
      TabIndex        =   5
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearchMonths 
      Caption         =   "Find monthly profits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5040
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdProfit 
      Caption         =   "Determine Profit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2400
      TabIndex        =   3
      Top             =   8280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   5535
      Left            =   480
      ScaleHeight     =   5475
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   2040
      Width           =   6495
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label lblWorkStation 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "the work station"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   14055
   End
End
Attribute VB_Name = "frmWorkStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'This program receives data from a file and sorts it into arrays
    'It also performs a calculation from a array data to make another array and then searchs that array
    

Private Sub cmdOpen_Click()
    'Declare variables
    Dim J As Integer
    
    'Prepare the file to be opened
    Open App.Path & "\bananas.txt" For Input As #1
    CTR = 0
    
    'Open file with a Do While Loop to sort it into arrays
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, bushels(CTR), cost(CTR), price(CTR), month(CTR)
    Loop
    
    Close #1
    
    'Print arrays neatly into a table using a for/next loop
    picResults.Print "Month"; Tab(20); "Thousands of Bushels", "Cost to produce", "Market price"
    picResults.Print "****************************************************************************************************************"
    For J = 1 To CTR
        picResults.Print month(J); Tab(20); bushels(J), , FormatCurrency(cost(J)), , FormatCurrency(price(J))
    Next J
    
    'Hide this button and make the next one visible
    cmdProfit.Visible = True
    cmdOpen.Visible = False
End Sub

Private Sub cmdProfit_Click()
    'Declare variables
    Dim J As Integer, Sum As Single, tProfit As Single, tBushels As Single
    
    'For next loop adds sum of bushels and determines total profit
    For J = 1 To CTR
        profit(J) = (price(J) - cost(J)) * (bushels(J) * 1000)
        tBushels = tBushels + bushels(J)
        tProfit = tProfit + profit(J)
    Next J
    
    'Print results
    picResults.Print "*****************************************************************************************************************"
    picResults.Print "Total"; Tab(20); tBushels
    MsgBox ("The net profit from this year's banana sales is " & FormatCurrency(tProfit))
    
    'Make next button visible
    cmdSearchMonths.Visible = True
    cmdProfit.Visible = False
End Sub


Private Sub cmdSearchMonths_Click()
    'Declare variables
    Dim J As Integer, bigProfit As Long, bigMonth As String
    
    bigProfit = 0
    picResults.Cls
    
    'Print monthly profits from new array made in previous button
    picResults.Print "Month"; Tab(20); "Profit"
    picResults.Print "***************************************************************************************************"
    For J = 1 To CTR
        picResults.Print month(J); Tab(20); FormatCurrency(profit(J))
    Next J
    
    'For next searchs profit array for largest value
    For J = 1 To CTR
        If profit(J) > bigProfit Then
            bigProfit = profit(J)
            bigMonth = month(J)
        End If
    Next J
    
    'Tell user what the most profitable month was
    MsgBox ("The most profitable month was " & bigMonth & " in which the company made " & FormatCurrency(bigProfit))
    
    'Allow user to see a button to switch to a new form
    cmdPromotion.Visible = True
End Sub

Private Sub cmdStop_Click()
    End
End Sub

Private Sub cmdSwitchBack_Click()
    'Allow user to go back to welcome room
    frmWelcome.Show
    frmWorkStation.Hide
End Sub

Private Sub cmdPromotion_Click()
    'Switch forms
    frmPromotion.Show
    frmWorkStation.Hide
End Sub
