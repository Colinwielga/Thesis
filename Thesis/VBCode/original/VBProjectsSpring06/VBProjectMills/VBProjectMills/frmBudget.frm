VERSION 5.00
Begin VB.Form frmBudget 
   Caption         =   "Bryan Mills"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   Picture         =   "frmBudget.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H80000012&
      Caption         =   "Find what you can afford to buy"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label lblBudget 
      BackColor       =   &H8000000D&
      Caption         =   "How much are you willing to spend?"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form asks for how much money you are willing to spend on your bass fishing trip(final.vbd)
'budget form (budget.frm)
'Bryan Mills
'March 24, 2006
Option Explicit
Private Sub cmdBudget_Click()
    Dim counter As Single
    budget = Val(InputBox("How much are you willing to spend?", "Budget"))
        'this takes an input box and saves the value under variable budget
            If budget > 500 Then
                MsgBox "Big Spender!", , "Hot Shot"
                'this outputs a quote regarding the value entered into the message box
            End If
        If budget >= 250 Then
            MsgBox "You have some spending to do...", , "Lets get going"
        End If
        'this if statement looked at the budget value and compared it to prices
        If budget >= 50 Then
            MsgBox "Lets buy some stuff...", , "Let's go to the store"
            Else: MsgBox "Get a job to get more money!", , "Cheap Skate"
        End If
            counter = 1
    Open App.Path & "\gear.txt" For Input As #1 'opens the file so it can be read
        Do Until EOF(1)
            Input #1, gear(counter), price(counter)
            counter = counter + 1
            'this reads the find stored as a parallel array so it can be accessed later
        Loop
    Close #1 'closes the file when done reading the array so other files can be read
    frmBudget.Hide
    frmGear.Show
    frmGear.loadbudget
    'this moves to the next form and loads the budget at the same time
End Sub

