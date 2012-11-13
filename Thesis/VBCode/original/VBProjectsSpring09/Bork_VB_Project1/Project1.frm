VERSION 5.00
Begin VB.Form frmName 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Customer Name"
   ClientHeight    =   5760
   ClientLeft      =   4560
   ClientTop       =   2955
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   Picture         =   "Project1.frx":0000
   ScaleHeight     =   101.6
   ScaleMode       =   0  'User
   ScaleWidth      =   111.389
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous Customer "
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton cmdCurrent 
      Caption         =   "Current Customer "
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Customer "
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Are you a new, previous, or current customer?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Wilderness Outfitters! "
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'The objective of this program is to provide an easy, user-friendly, interactive
'method to organize camping trips for Wilderness Outfitters.  It is not intended
'to be implemented nor does it have the capability, however it is an idea of
'what could work for this type of business.
'
'frmName
'The purpose of this form is to aquire the customer's name and the current year
'in order to create a directory as well as a text file for each customer.

Private Sub cmdCurrent_Click()
    user = InputBox("Enter first and last name", "Customer Name")
    year = InputBox("Enter current year please!", "Year")
    
    'Takes name of current user and reads their file into an array.
    Open App.Path & "\Customers\" & user & "\" & user & year & ".txt" For Input As #1
        Do Until EOF(1)
            counter = counter + 1
            Input #1, Items(counter), Prices(counter), Requests(counter), Subtotals(counter)
        Loop
    Close #1
    
    frmName.Hide
    frmStartup.Show
    frmDisplay.Show
    
End Sub

Private Sub cmdNew_Click()
    'This takes in the name of the new user and creates a new directory and text
    'file for the user.
    'MkDir creates the specified directory while simply opening a non-existant
    'file creates the file.
    user = InputBox("Enter First and Last name please!", "New User")
    year = InputBox("Enter current year please!", "Year")
    
    MkDir App.Path & "\Customers\" & user & ""
    
    Open App.Path & "\customer.txt" For Input As #1
        Do Until EOF(1)
            counter = counter + 1
            Input #1, Items(counter), Prices(counter), Requests(counter), Subtotals(counter)
        Loop
    Close #1
    
    Open App.Path & "\Customers\" & user & "\" & user & year & ".txt" For Output As #1
        For pos = 1 To counter
            Write #1, Items(pos), Prices(pos), Requests(pos), Subtotals(pos)
        Next pos
    Close #1
    
    
    MsgBox "Congratulations " & user & " , you are officially a member of the Wilderness Outfitters family!"
    
    frmName.Hide
    frmStartup.Show
    frmDisplay.Show
    
End Sub

Private Sub cmdPrevious_Click()
    'Again it takes the user's name and the year but because this is a previous
    'customer the program opens their directory and creates a new text file.
    user = InputBox("Enter first and last name", "Customer Name")
    year = InputBox("Enter current year please!", "Year")
    
    Open App.Path & "\customer.txt" For Input As #1
        Do Until EOF(1)
            counter = counter + 1
            Input #1, Items(counter), Prices(counter), Requests(counter), Subtotals(counter)
        Loop
    Close #1
    
    Open App.Path & "\Customers\" & user & "\" & user & year & ".txt" For Output As #1
        For pos = 1 To counter
            Write #1, Items(pos), Prices(pos), Requests(pos), Subtotals(pos)
        Next pos
    Close #1
    
    MsgBox ("Hello " & user & ", and welcome back to Wilderness Outfitters!")
    
    frmName.Hide
    frmStartup.Show
    frmDisplay.Show
End Sub
