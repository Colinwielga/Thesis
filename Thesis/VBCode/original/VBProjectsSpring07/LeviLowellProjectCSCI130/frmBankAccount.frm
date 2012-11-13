VERSION 5.00
Begin VB.Form frmBankAccount 
   BackColor       =   &H00000000&
   Caption         =   "Your Bank Account"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000A&
      Height          =   3495
      Left            =   3240
      ScaleHeight     =   3435
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000C000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   3495
   End
   Begin VB.CommandButton cmdInputbalance 
      BackColor       =   &H0000C000&
      Caption         =   "Input your current account balance"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Label lblBalance 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Watch your balance, don't overdraw!"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Image ImgBankAccount 
      BorderStyle     =   1  'Fixed Single
      Height          =   4860
      Left            =   120
      Picture         =   "frmBankAccount.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2985
   End
   Begin VB.Label lblBankAccount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmBankAccount.frx":5E89
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2055
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmBankAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This application provides the user with the ability to maintain a balance, input transactions and deductions.
'The program keeps track of the balance and displays the balance and the transaction.  If the user overdraws
'their balance the program will warn them of this fact.

Private Sub cmdClear_Click()
picResults.Cls      'Clears the picturebox
End Sub

Private Sub cmdInputbalance_Click()

Dim X As Single          'Sets the variables
Dim sum As Single

sum = 0                 'Initiates the variables
X = 0

picResults.Print "Your Bank Account"        'Prints the message into the picturebox
X = InputBox("Enter an amount, or -999 to indicate the end of your data")       'Asks the users to input their initial balance
If X < 0 Then
    X = InputBox("Invalid Input, please re-enter", , "warning")     'Asks the user to re-enter when they're varialbe is invalid
End If
        
Do While X <> -999      'loops through InputBoxes, the user can enter as many variables as necessary, when the flag -999 is entered the loop ends
        sum = sum + X       'Adds the new input to the existing sum
            picResults.Print FormatCurrency(X, 2)       'Prints the amount entered into the picturebox
            picResults.Print "new balance: " & FormatCurrency(sum, 2)       'Prints the new balance into the picturebox
            X = InputBox("Enter an amount, or -999 to indicate the end of your data")       'asks the user for a new input
Loop
    MsgBox " Your final account balance is " & FormatCurrency(sum, 2), , "Final Balance"        'Displays the final account balance if the balance is positive
      If sum < 0 Then
        MsgBox "Warning Your Balance is overdrawn! " & FormatCurrency(sum, 2), , "Overdrawn"        'If the account is overdrawn, less than zero, a message box is displayed warning the user
      End If        'Ends the if statement
End Sub

Private Sub cmdReturn_Click()
frmBankAccount.Hide     'Hides frmBankAccount
FrmMain.Show        'Shows frmMain
End Sub

