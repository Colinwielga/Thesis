VERSION 5.00
Begin VB.Form example 
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   Picture         =   "Example.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt4 
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Text            =   "0"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txt3 
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txt2 
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdinterestrate 
      Caption         =   "Click me to see how many number of period you need so you can save the amount of money you wish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      TabIndex        =   3
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Calculate Present Value"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Calculate Future Value"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lbl4 
      BackColor       =   &H80000009&
      Caption         =   "<----Number of period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lbl6 
      BackColor       =   &H80000013&
      Caption         =   "Interest rate is given at 5%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   13
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label lbl3 
      BackColor       =   &H80000009&
      Caption         =   "Money you have recevied-->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label lbl2 
      BackColor       =   &H80000009&
      Caption         =   "<----Number of period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   11
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lbl1 
      BackColor       =   &H80000009&
      Caption         =   "Money you have today ----->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lbl5 
      BackColor       =   &H80000013&
      Caption         =   "Interest rate is given at 5%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblexample 
      BackColor       =   &H80000012&
      Caption         =   "Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Compound Interest
'Form:example
'Author:Yik Pang Ngan (Banny)
'Date Written:Oct 9 2009
Option Explicit
'this form will show experiments how the formulas work.
Dim PV As Single, period As Integer, interest As Single, FV As Single
'define all the varibles I will be using


Private Sub cmd1_Click()
'define the text boxes to each varible
PV = txt1.Text
period = txt2.Text
interest = 0.05

FV = PV * (1 + interest) ^ period 'this is the formula I use to calculate the future value

MsgBox ("You will have " & FormatCurrency(FV, 2) & " after " & period & " years if interest rate is 5%")
'this will tell the result based on the information from the text boxes
End Sub

Private Sub cmd2_Click()

'define the text boxes to each varible
FV = txt3.Text
period = txt4.Text
interest = 0.05

PV = FV / (1 + interest) ^ period 'this is the formula I use to calculate the future value

MsgBox (period & " years ago, you had " & FormatCurrency(PV, 2) & " if the interest is 5%")
'this will tell the result based on the information from the text boxes
End Sub

Private Sub cmdinterestrate_Click()

'this will tell you how many periods required to get Future Value from Present Value and the interest rate
PV = InputBox("Please enter how much money you have now.") 'this can input the varibles by using inputbox feature
FV = InputBox("Enter how much money you need to save.")
interest = InputBox("what is the interest rate")


period = Log(FV) - Log(PV) / Log(1 + interest)


If FV > PV Then
    MsgBox ("You will need " & period & " years in order to save $" & FV)

ElseIf PV >= FV Then
    MsgBox ("You already have more than the money you wish to save.")
'if present value is greater than future value, so there is no need to save money.

End If

End Sub



Private Sub cmdQuit_Click()
End
End Sub
