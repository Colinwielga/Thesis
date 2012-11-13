VERSION 5.00
Begin VB.Form frmLoanShark 
   Caption         =   "Get Money"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   7935
      Left            =   0
      Picture         =   "frmLoanShark.frx":0000
      ScaleHeight     =   7875
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.PictureBox picResults 
         BackColor       =   &H00808000&
         Height          =   495
         Left            =   3600
         ScaleHeight     =   435
         ScaleWidth      =   4155
         TabIndex        =   4
         Top             =   6960
         Width           =   4215
      End
      Begin VB.CommandButton cmdBack2 
         BackColor       =   &H00404080&
         Caption         =   "Go Back to Casino"
         Height          =   855
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5880
         Width           =   2295
      End
      Begin VB.CommandButton cmdATM 
         BackColor       =   &H0000C0C0&
         Caption         =   "Go to local ATM"
         Height          =   855
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5880
         Width           =   2415
      End
      Begin VB.CommandButton cmdLoan 
         BackColor       =   &H0000C000&
         Caption         =   "Ask the Loan Shark"
         Height          =   855
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5880
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmLoanShark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Mystake Lake Casino
'Authors: David Johnson And Jeremy Iverson
'Date: Monday, November 5, 2007

Option Explicit
'This form allows the user to get money to enter the casino

Private Sub cmdATM_Click()
    ''This button allows the user to make 1 withdrawal up to $100
    Dim balanceatm As Single, amount As Single, nameatm As String
    nameatm = InputBox("Please enter your name for identification purposes.", "Name")
    If LCase(nameatm) = LCase(nameglobal) Then
        amount = InputBox("Your current balance is $100. How much do you want to withdraw?", "Withdrawl")
        balanceatm = 100 - amount
        balanceglobal = balanceglobal + amount
        If balanceatm < 0 Then
            MsgBox "Error, your withdrawl exceeds your balance.", , "Error"
            balanceglobal = 0
        Else
            MsgBox "Your new balance is " & FormatCurrency(balanceatm) & ". Have a nice day!", , "New Balance"
            picResults.Print "Your wallet holds " & FormatCurrency(balanceglobal) & "."
    
            'After you leave the ATM, it gets robbed and displays the time of robbery
    
            MsgBox "ATM is robbed as soon as you leave, last transaction made at " & Time
            cmdATM.Enabled = False
        End If
    Else
        MsgBox "Are you trying to use someone else's account", , "Just Wondering"
    End If
    
End Sub

Private Sub cmdBack2_Click()
    'Returns the user to outside of the casino
    frmLoanShark.Hide
    frmCasino.Show
End Sub

Private Sub cmdLoan_Click()
    Dim balanceloan As Single, nameloan As String
    
    clicked = True 'When leaving casino you have to pay them back
    
    nameloan = InputBox("Alright, I talk, you listen. First I need a name, what is it?", "Name")
    
    'Depending on how much money you loan, they say something different
    'The user can only take out $10,000 or less
    If balanceglobal < 10000 Then
        If LCase(nameloan) = LCase(nameglobal) Then
            balanceloan = InputBox("Number Two, The limit is $10,000. How much will it be tonight?", "Amount")
            If 10000 >= balanceglobal + balanceloan Then
                balanceglobal = balanceglobal + balanceloan
                Select Case balanceloan
                    Case 8000 To 10000
                        MsgBox "Big Spender for a new client. You have 24 hours.  Good luck.", , "Loan Shark"
                    Case 4000 To 8000
                        MsgBox "Looking for some fun 'eh. You don't get enough at home do ya?! You have 24 hours.", , "Loan Shark"
                    Case 500 To 4000
                        MsgBox "You know what I learned in high school: the higher the risk, the higher return. Well, You have 24 hours.", , "Loan Shark"
                    Case 0.01 To 500
                        MsgBox "Sounds good, you have 24 hours.", , "Loan Shark"
                    Case Else
                        MsgBox "Are you homeless! This is a business.  Get outta here before we gamble on you.", , "Loan Shark"
                End Select
            Else
                MsgBox "Should we refresh your memory about the $10,000 limit???", , "Big Mistake"
                balanceglobal = 0
            End If
        Else
            MsgBox "Get your sorry behind outta here before we make it really sorry!", , "What is your real name"
            frmLoanShark.Hide
            frmCasino.Show
        End If
    Else
        MsgBox "We got nothing more for you!", , "Scram"
    End If
    'This prevents the user from taking out more than $10,000
    If balanceloan > 10000 Or balanceloan <= 0 Then
        balanceloan = 0
    End If
    temp = balanceglobal
    picResults.Cls
    'Displays the user's current balance in a picturebox
    picResults.Print "Your wallet holds " & FormatCurrency(balanceglobal) & "."
End Sub

