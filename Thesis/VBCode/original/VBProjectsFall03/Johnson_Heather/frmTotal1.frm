VERSION 5.00
Begin VB.Form frmTotal1 
   BackColor       =   &H00FF80FF&
   Caption         =   "Total"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FF8080&
      Caption         =   "QUIT"
      Height          =   1095
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Cmdtotal 
      BackColor       =   &H0080FF80&
      Caption         =   "FINAL TOTAL!!!!!!"
      Height          =   1215
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.PictureBox picendtotal 
      BackColor       =   &H00FFC0FF&
      Height          =   6855
      Left            =   5520
      ScaleHeight     =   6795
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H00FF80FF&
      Caption         =   "HERE IS YOUR TOTAL FOR ALL YOUR ORDERS!!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmTotal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Cheerleading (Cheerleading.vbp)
'Form Name : Total (Total1.frm)
'Author: Heather Johnson
'Date Written: October 28, 2003
'Purpose of Form:'give you the total of all your Orders
                 'this form will keep track of all your previous orders and then carry
                 'the total to the picture box on this form
Option Explicit
Dim FinalTotal As Single
Private Sub cmdquit_Click()
MsgBox "THANKS FOR YOUR ORDER, HAVE A WONDERFUL DAY!!!!!!", , "Thanks"
    'when you push the wuit button you will get a little messgae box thanking the orderer
End 'ends the program
End Sub

Private Sub Cmdtotal_Click()
FinalTotal = 0
picendtotal.Print "Your Uniform Body Color is "; Tab(42); UniColor
picendtotal.Print "Your Accent Color is "; Tab(42); AccColor
picendtotal.Print "Your total for the shells is"; Tab(42); FormatCurrency(TotalCost, 2)
    'prints out the total after the discount from the shells form
picendtotal.Print "Your Total for the lettering is"; Tab(42); FormatCurrency(TotalCostLettering, 2)
    'prints out the total after the discount from the lettering form
picendtotal.Print "Your Total for the skirts is"; Tab(42); FormatCurrency(TotalSkirts, 2)
FinalTotal = TotalCost + TotalCostLettering + TotalSkirts
    'adds the total from the shells, the lettering, and the skirts and that is the final total cost
picendtotal.Print "*********************************************************************************************************************************"
picendtotal.Print "Your Final Total is"; Tab(42); FormatCurrency(FinalTotal, 2)
    'prints out the final total cost for your uniforms
End Sub


