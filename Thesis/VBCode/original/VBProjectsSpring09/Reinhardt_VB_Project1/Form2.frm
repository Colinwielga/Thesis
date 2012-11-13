VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   7335
   ClientLeft      =   10140
   ClientTop       =   3075
   ClientWidth     =   8460
   LinkTopic       =   "Form2"
   ScaleHeight     =   7335
   ScaleWidth      =   8460
   Begin VB.CommandButton cmderase 
      BackColor       =   &H0000FFFF&
      Caption         =   "Clear Results"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdfinish 
      BackColor       =   &H000000FF&
      Caption         =   "Go back to Main Form"
      Height          =   975
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Txteuro 
      Height          =   975
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdeuro 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exchange Euro to Dollars"
      Enabled         =   0   'False
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox picresult 
      BackColor       =   &H8000000E&
      Height          =   1695
      Left            =   2880
      ScaleHeight     =   1635
      ScaleWidth      =   4635
      TabIndex        =   2
      Top             =   4200
      Width           =   4695
   End
   Begin VB.TextBox txtInput 
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdexchange 
      BackColor       =   &H000000FF&
      Caption         =   "Exchange Dollars to Euro"
      Enabled         =   0   'False
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Enter Euro Amount Here"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Enter Dollar Amout Here"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Information for German/Austria tourists
'Form 2
'Joseph Reinhardt
'March 20, 2009
'To recieve user-input and calculate the amount of opposite currency that is
'This form is only accesible from the first form

'Calculate amount of euro from american dollars
Private Sub cmdexchange_Click()
'Declare Variables
Dim Money As Single, exchange As Single

'assign the user-input from the Textbox to a variable
Money = txtInput.Text

'Perform the calculation assigning answer to variable
exchange = Money * 0.738062

'Print results
picresult.Print FormatCurrency(Money) & " exchanges to " & "€" & FormatNumber(exchange, 2)

End Sub
'calcualte amount of dollar from euro
Private Sub cmdeuro_Click()
'Declare Variables
Dim Geld As Single, Change As Single

'Assign the user-input from the Textbox to a variable
Geld = Txteuro.Text

'Perform The calcualtion assigning the answer to a variable
Change = Geld * 1.35489

'Print Results
picresult.Print "€" & FormatNumber(Geld) & " exchanges to " & FormatCurrency(Change)

End Sub
'end form
Private Sub cmdfinish_Click()

'Show Form 1 again
Form1.Show

End Sub
'Clear picture box
Private Sub cmderase_Click()

'Clear picresult picture box
picresult.Cls

End Sub
'disallow use of cmdeuro until user-input is obtained
Private Sub Txteuro_Change()

'Once user input is obtained, cmdeuro button is enabled
cmdeuro.Enabled = True

End Sub
'disallow the use of cmdchange until user-input is obtained
Private Sub txtInput_Change()

'Once user input is obtained, cmdchange button is enabled
cmdexchange.Enabled = True

End Sub
