VERSION 5.00
Begin VB.Form frmorder 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Front Page"
      Height          =   735
      Left            =   960
      TabIndex        =   8
      Top             =   6960
      Width           =   4935
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Find My Total"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   6120
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      Height          =   4695
      Left            =   2880
      ScaleHeight     =   4635
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton cmdhat 
      Caption         =   "Buy a Hat"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdposter 
      Caption         =   "Buy a Poster"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdjacket 
      Caption         =   "Buy a Jacket"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdshirt 
      Caption         =   "Buy a Tee-Shirt"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdsweat 
      Caption         =   "Buy a Sweatshirt"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdmerchandise 
      Caption         =   "View The Merchandise"
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lbldeal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If you buy more than $50 worth of merchandise you get $5 off and if you buy $100 worth of merchandise you get $15 off!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -1680
      Picture         =   "frmorder.frx":0000
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   12000
   End
End
Attribute VB_Name = "frmorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Vikings Fan Page

'Form Name: Order

'Written by Jim Zrust

'Date: November 1, 2006

'Form Objective:this form functioned basically as the cash register for the team store where the user was able
'to decide what products they wanted to buy and how many of the product. i decided to offer a
'discount in order to do more detailed coding

Option Explicit
Dim runningtotal As Double 'runningtotal is a variable that i plan on using throughout the form

Private Sub cmdhat_Click() 'by clicking on the button it automatically prints to the picturebox and will add the cost of the item to the runningtotal variable
Dim hatcost As Single
hatcost = 50
picResults.Print "Hat:", , FormatCurrency(hatcost, 2)
runningtotal = runningtotal + hatcost
End Sub

Private Sub cmdjacket_Click() 'see cmdhat
Dim jacketcost As Single
jacketcost = 100
picResults.Print "Jacket:", , FormatCurrency(jacketcost, 2)
runningtotal = runningtotal + jacketcost
End Sub

Private Sub cmdmerchandise_Click() 'allows the user to view the merchandise
frmorder.Hide
frmmerchandise.Show
End Sub

Private Sub cmdposter_Click() 'see cmdhat
Dim postercost As Single
postercost = 5
picResults.Print "Poster:", , FormatCurrency(postercost, 2)
runningtotal = runningtotal + postercost
End Sub

Private Sub cmdReturn_Click() 'allows the user to return to the front page
frmorder.Hide
frmhome.Show
End Sub

Private Sub cmdshirt_Click() 'see cmdhat
Dim teeshirtcost As Single
teeshirtcost = 17
picResults.Print "Tee-Shirt:", , FormatCurrency(teeshirtcost, 2)
runningtotal = runningtotal + teeshirtcost
End Sub

Private Sub cmdsweat_Click() 'see cmdhat
Dim Sweatshirtcost As Single
Sweatshirtcost = 50
picResults.Print "Sweatshirt:", , FormatCurrency(Sweatshirtcost, 2)
runningtotal = runningtotal + Sweatshirtcost
End Sub

Private Sub cmdtotal_Click()
Dim Tax As Single
Dim Total As Single
Total = runningtotal
picResults.Print , "-------------"
Select Case Total 'select case statement was used to figure out whether a discount should be subtracted from the total
    Case Is > 100
        Total = Total - 15
        picResults.Print "Your Total is:", , FormatCurrency(Total, 2)
    Case 50 To 100
        Total = Total - 5
        picResults.Print "Your Total is:", , FormatCurrency(Total, 2)
    Case Else
        picResults.Print "Your Total is:", , FormatCurrency(Total, 2)
End Select
End Sub

