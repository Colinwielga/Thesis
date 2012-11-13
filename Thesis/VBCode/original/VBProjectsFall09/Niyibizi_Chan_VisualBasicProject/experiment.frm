VERSION 5.00
Begin VB.Form experiment 
   BackColor       =   &H00004080&
   Caption         =   "Form2"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form2"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo2IncomeStatement 
      BackColor       =   &H00008000&
      Caption         =   "Click to explore Income statement"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9240
      Width           =   3495
   End
   Begin VB.TextBox lblRevenues 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   7
      Text            =   "                 Profit and Loss "
      Top             =   480
      Width           =   4815
   End
   Begin VB.PictureBox picResults2 
      FillColor       =   &H80000014&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   8040
      ScaleHeight     =   4275
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   3960
      Width           =   5655
   End
   Begin VB.CommandButton cmdcompute 
      BackColor       =   &H00404080&
      Caption         =   "Let see how you did this year. (Gain or loss)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   4695
   End
   Begin VB.TextBox txtexpense 
      Height          =   855
      Left            =   3000
      TabIndex        =   4
      Text            =   "0"
      Top             =   6480
      Width           =   2055
   End
   Begin VB.TextBox txtrevenue 
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Text            =   "0"
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label lblexpense 
      BackColor       =   &H000080FF&
      Caption         =   "Enter the expenses you used this year. ------------->"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label lblownbusiness 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"experiment.frx":0000
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Label lblRevenue 
      BackColor       =   &H0000C0C0&
      Caption         =   "Enter the revenue you earned this year.  --------->"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   1935
   End
End
Attribute VB_Name = "experiment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Accounting basics and Income statement
'Form 2:Experiment with profit and loss
'Author:Patrick Niyibizi and Frankie Chan
'Date Written:September 30th 2009
'Objective:To explore the basis of profit or loss that is key to the income statement that will be introduced in the next form.
Option Explicit

Private Sub cmdCompute_Click()

Dim revenue As Single, expenses As Single, Gain As Single   'Declare variables to be used

    revenue = txtrevenue      'Assign the variables to textboxes
    expenses = txtexpense
    
    Gain = revenue - expenses     'Calculate the outcome
    
    If Gain >= 0 Then
        picResults2.Picture = LoadPicture(App.Path & "\Images\profit.jpg")                 'If you had a profit, display the profit as well as an appropriate picture
        MsgBox ("You have made " & FormatCurrency(Gain, 2) & " this year. Awesome")
    ElseIf Gain < 0 Then
        picResults2.Picture = LoadPicture(App.Path & "\Images\loss.jpg")                    'If you had a loss,display the loss as well as an appropriate picture
        MsgBox ("You have lost " & FormatCurrency(Gain, 2) & " this year. You need to do better than that.")
    End If
    
End Sub

Private Sub cmdGo2IncomeStatement_Click()      'Go to the third form
    experiment.Hide
    IncomeStatfrm.Show
    
    
End Sub
