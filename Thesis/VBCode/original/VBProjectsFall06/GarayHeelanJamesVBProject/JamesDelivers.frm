VERSION 5.00
Begin VB.Form frmFrontPage 
   Caption         =   "Welcome to James Delivers!"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   Picture         =   "James Delivers.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   4920
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   11055
      Left            =   -120
      Picture         =   "James Delivers.frx":1CF6E
      ScaleHeight     =   10995
      ScaleWidth      =   13995
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.CommandButton cmdLogOut 
         Caption         =   "Log Out"
         Height          =   735
         Left            =   600
         TabIndex        =   5
         Top             =   6600
         Width           =   3135
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   735
         Left            =   600
         TabIndex        =   4
         Top             =   5640
         Width           =   3135
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register!"
         Height          =   855
         Left            =   4560
         TabIndex        =   3
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox txtLogin 
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   4080
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmFrontPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmFrontPage
'Written by James Garay Heelan
'on 11-2-06
'The purpose of this program is to allow a user to register with and order from
'the online grocery delivery service, "James Delivers".  Given the popularity
'of online grocer Simon Delivers, I felt that it was time to give them a
'bit of competition.

'This form is the first form the user sees.  It welcomes the user and allows
'him or her to either login with an existing account, or to register as
'a new user.

Option Explicit
Private Sub cmdLogin_Click()
Dim Found As Boolean
    Sum = 0
    Size = 0
    Pos = 0
    Open App.Path & "/RegisteredUsers.txt" For Input As #2 'opens the registered user database
    Do Until EOF(2) 'commands that the sort be performed till the end of the file
        Size = Size + 1 'increases the recorded size of the file by 1
        Input #2, Names(Size), Address(Size), City(Size), State(Size), Zip(Size), PaymentMethod(Size), CreditCardNumber(Size), ExpirationDate(Size), RLogin(Size), RPassword(Size) 'reads and organizes the information of all registered James Delivers users
        Loop 'Repeats the process from the beginning of the sort
    Close #2 'closes the file
            
    For Pos = 1 To Size 'Declares the perameters of the search to look through as many entries as the size of the file
        If (txtLogin(0).Text = RLogin(Pos)) And (txtPassword(1).Text = RPassword(Pos)) Then 'searches for a match between the entered login and password with a login and password from the registered users file
            Found = True 'if a match is found, then a variable indicates found
            PurchaseCode = Pos 'the user number is recorded for future use in the program
        End If
    Next Pos
    
    If Found = False Then 'if the login and password do not match, then
        MsgBox "Please register with us!", , "Login or Password not recognized" 'a messagebox is displayed asking the user to register
    Else
        frmFrontPage.Hide 'if match is found, the frontpage is hidden and
        frmGroceryStore.Show 'the central shopping menu is displayed to the user
        Open App.Path & "/PurchasedItems.txt" For Append As #10 'a shopping cart file is either created or opened if it already exists, the file must be created if it does not exists, otherwise the clearing function will not work and the program will jam
            Close #10 'the file is closed
        Kill App.Path & "/PurchasedItems.txt" 'the shopping cart file is deleted, to get rid of any previous orders by previous users
          
    End If
    
End Sub

Private Sub cmdLogOut_Click()
    End 'exits the program
End Sub

Private Sub cmdRegister_Click()
    frmFrontPage.Hide 'the frontpage is hidden
    frmRegister.Show 'the registration page is displayed for the user
End Sub


