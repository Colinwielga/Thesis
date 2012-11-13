VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14505
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H80000014&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H8000000E&
      Height          =   8175
      Left            =   3120
      ScaleHeight     =   8115
      ScaleWidth      =   10155
      TabIndex        =   7
      Top             =   1440
      Width           =   10215
   End
   Begin VB.CommandButton cmddisplay 
      BackColor       =   &H80000014&
      Caption         =   "Display the Financial Statements"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H80000014&
      Caption         =   "Compute the Rate of Return on Assets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdprofit 
      BackColor       =   &H80000014&
      Caption         =   "Compute the Profit Margin on Sale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdquick 
      BackColor       =   &H80000014&
      Caption         =   "Compute the Quick Ratio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdcurrent 
      BackColor       =   &H80000014&
      Caption         =   "Compute the Current Ratio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdload 
      BackColor       =   &H80000014&
      Caption         =   "Upload the Financial Statements"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bridget Spaniol"
      Height          =   255
      Left            =   12480
      TabIndex        =   9
      Top             =   11280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "International Paper 2003 Financial Analysis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   9285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim title(1 To 100) As String
Dim amount(1 To 100) As Single
Dim J As Integer
Dim ctr As Integer
Dim currentliabilities As Single
Dim currentassets As Single
Dim current As Single
Dim netreceivables As Single
Dim cash As Single
Dim quick As Single
Dim NetIncome As Single
Dim NetSales As Single
Dim Profit As Single
Dim TotalAssets As Single
Dim Rate As Single

'International Paper financial interpretation (Financial Analysis)
'Form 1 (project 1)
'Bridget Spaniol
'3/13/04
'This form is used to interprete the financial statements and use them to calculate useful information to
'determine the financial wellbeing of Internation Paper.

Private Sub cmdcurrent_Click() 'This calculates the current ratio
picresults.Cls
For J = 1 To 100
    If title(J) = "Total Current Assets" Then
        currentassets = amount(J)
    End If
Next J
For J = 1 To 100
    If title(J) = "Total Current Liabilities" Then
        currentliabilities = amount(J)
    End If
Next J
current = currentassets / currentliabilities
picresults.Print "The Current Ratio is", current, "."
picresults.Print
picresults.Print "The current ratio measures the short-term debt-paying ability of International Paper."
picresults.Print "The company is sitting very well with a ratio of 1.3724, meaning that the amount of liabilities the "
picresults.Print "company has can be covered by the assets a total of 1.37 times."

End Sub

Private Sub cmddisplay_Click() 'This displays the financial information.
picresults.Cls
picresults.Print "Balance Sheet  (in millions of dollars)"
picresults.Print
picresults.Print "Current Assets"
For J = 1 To 8
    picresults.Print title(J), Tab(30), FormatCurrency(amount(J), 2)
Next J
picresults.Print
picresults.Print "Current Liabilities"
For J = 9 To 15
    picresults.Print title(J), Tab(30), FormatCurrency(amount(J), 2)
Next J
picresults.Print
picresults.Print "Income Statement Information"
picresults.Print
For J = 16 To 18
    picresults.Print title(J), Tab(30), FormatCurrency(amount(J), 2)
Next J
End Sub

Private Sub cmdload_Click() 'This inputs the information into arrays.
Open "N:\CS130\handin\Spaniol, Bridget\financial.txt" For Input As #1
    ctr = 0
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, title(ctr), amount(ctr)
Loop
cmdload.Visible = False
cmdcurrent.Visible = True
cmdquick.Visible = True
cmdprofit.Visible = True
cmdreturn.Visible = True
cmddisplay.Visible = True

End Sub

Private Sub cmdprofit_Click() 'This calculates the profit margin on sales.
picresults.Cls
For J = 1 To 100
    If title(J) = "Total Net Income" Then
        NetIncome = amount(J)
    End If
Next J
For J = 1 To 100
    If title(J) = "Net Sales" Then
        NetSales = amount(J)
    End If
Next J
Profit = (NetIncome / NetSales) * 100
picresults.Print "The Profit Margin on sales is", Profit, "%."
picresults.Print
picresults.Print "The Profit Margin on sales measures International Paper's profit per dollar of sales."


End Sub

Private Sub cmdquick_Click() 'This calculates the quick ratio.
picresults.Cls
For J = 1 To 100
    If title(J) = "Cash" Then
        cash = amount(J)
    End If
Next J
For J = 1 To 100
    If title(J) = "Net receivables" Then
        netreceivables = amount(J)
    End If
Next J
For J = 1 To 100
    If title(J) = "Total Current Liabilities" Then
        currentliabilities = amount(J)
    End If
Next J
quick = (cash + netreceivables) / currentliabilities
picresults.Print "The quick ratio is", quick, "."
picresults.Print
picresults.Print "The quick ratio states the ability of International Paper to quickly turn assets into cash to cover "
picresults.Print "any unforseen problems in the company.  The quick ratio should be above 1 but in comparison"
picresults.Print "to other manufacturing companies, a .3 is the relative average."

End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub cmdreturn_Click() 'This calculates the rate of return on assets.
picresults.Cls
For J = 1 To 100
    If title(J) = "Total Net Income" Then
        NetIncome = amount(J)
    End If
Next J
For J = 1 To 100
    If title(J) = "Total Assets" Then
        TotalAssets = amount(J)
    End If
Next J
Rate = (NetIncome / TotalAssets) * 100
picresults.Print "The Rate of Return on Assets is,"; Rate, "%."
picresults.Print
picresults.Print "The Rate of Return on Assets describes International Paper's ability to use assets to"
picresults.Print "generate earnings.  This rate is low for a manufacturing business suggesting in efficient use"
picresults.Print "of inventory supply."
End Sub

Private Sub Form_Load()
Form4.Show
End Sub
