VERSION 5.00
Begin VB.Form UsefulInformation 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Useful Information"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcost 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Average Cost of General Products  (In Australian Dollars)"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Back To Home Page"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox txtmoney 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdconvert 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Convert Now"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdcalc2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Calculate Celsius to Fahrenheit"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox txtCtoF 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Australia's Average Temperature"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C0C0FF&
      Height          =   4695
      Left            =   4320
      ScaleHeight     =   4635
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdcalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Calculate Fahrenheit to Celsius"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txttemp 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Ashley 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ashley K. Smithson"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Temperature Converter Enter Temperature in Celsius"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Convert Australian Dollar to U.S. Dollar.  Enter amount in Australian Dollar:"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Temperature Converter  Enter Temperature in Fahrenheit"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "UsefulInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Australia
'Form Name: UsefulInformation
'Author: Ashley Smithson
'Date: October 31, 2005
'Purpose of Form: To help with the difficult calculations for currency and temperature.
Option Explicit
Dim month(1 To 100) As String, low(1 To 100) As Integer, high(1 To 100) As Integer, CTR As Integer, J As Integer
Dim F As Single, C As Single, U As Single, S As Single
Dim product(1 To 100) As String, price(1 To 100) As String, D As Single

Private Sub cmdback_Click()
UsefulInformation.Hide
FinalProject2.Show
End Sub

Private Sub cmdcalc2_Click()
picresults.Cls
C = Val(txtCtoF.Text)
F = 9 / 5 * C + 32
picresults.Print C; "Degrees Celsius is:"; F; "Fahrenheit."
End Sub

Private Sub Text1_Change()
End Sub

Private Sub cmdcalc_Click()
picresults.Cls 'clears picture box
F = Val(txttemp.Text)
C = 5 / 9 * (F - 32)
picresults.Print F; "Degrees Fahrenheit is:"; C; "Celsius."
'calculates and prints results
End Sub

Private Sub cmdconvert_Click()
picresults.Cls
U = Val(txtmoney.Text)
S = U * 0.75
picresults.Print FormatCurrency(U); Tab(9); "Australian =";
picresults.Print Tab(20); FormatCurrency(S);
picresults.Print Tab(28); "American."
'calculates and prints results putting the words in specific placing due to tab
End Sub

Private Sub cmdcost_Click()
Open App.Path & "\goods.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, product(CTR), price(CTR)
Loop
Close #1
picresults.Cls
picresults.Print "Product"; Tab(30); "Price"
picresults.Print "***********************************************"
For D = 1 To CTR
  picresults.Print product(D); Tab(30); price(D)
Next D
'gets info displays
End Sub

Private Sub cmdfind_Click()
Open App.Path & "\temperature.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, month(CTR), low(CTR), high(CTR)
Loop
Close #1
picresults.Cls
picresults.Print "Month"; "Low Temp (C)"; Tab(25); "High Temp(C)"
picresults.Print "**********************************************************"
For J = 1 To CTR
  picresults.Print month(J), low(J), high(J)
Next J
'ditto
End Sub

