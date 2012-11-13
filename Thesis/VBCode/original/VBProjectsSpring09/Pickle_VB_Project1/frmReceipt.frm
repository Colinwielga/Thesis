VERSION 5.00
Begin VB.Form frmReceipt 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FF00FF&
      Caption         =   "Previous Screen"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   2535
   End
   Begin VB.PictureBox picWoman 
      Height          =   4455
      Left            =   9360
      Picture         =   "frmReceipt.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   4395
      TabIndex        =   5
      Top             =   4800
      Width           =   4455
   End
   Begin VB.PictureBox picMatt 
      Height          =   4455
      Left            =   7680
      Picture         =   "frmReceipt.frx":8CFB
      ScaleHeight     =   4395
      ScaleWidth      =   4395
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H00FF00FF&
      Caption         =   "Purchase Now!"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FF00FF&
      Caption         =   "View My Total"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   4695
      Left            =   1320
      ScaleHeight     =   4635
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblYou 
      BackColor       =   &H0000FF00&
      Caption         =   "This could be you!!!"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   12480
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Build Your Own Home Gym
'Form Name: frmReceipt
'Author: Michelle Pickle
'Date Written: March 12th 2009
'The purpose of this screen to allow the user to view their grand total, including tax, of the items they purchased.
    'The picutres encourage the people to buy in order to look like the individuals in the picture.
'requires the user to declare their variables before the program is functional
Option Explicit

Private Sub cmdPrevious_Click()
'allows the user to navigate to the previous form
    frmGyms.Visible = True
    frmReceipt.Visible = False
End Sub

Private Sub cmdPurchase_Click()
'this button is an essence purchasing the product, thanking the user, and reminding him/her of his/her total.
    MsgBox "Your total was " & FormatCurrency(runningtotal) & ". Thank you for your purchase!", , "Thank You"
    End
    
End Sub

Private Sub cmdQuit_Click()
'the program ends
    End
End Sub

Private Sub cmdTotal_Click()
'calcualtes the final total, including tax, of all the objects combined
'declares the variables
Dim tax As Double
Dim grandtotal As Double

'calculates the final total
'prints the subtotal(before tax), the tax, and then the final total (including tax)
    picResults.Print "Handheld Equipment"; Tab(30); FormatCurrency(handheldtotal)
    picResults.Print "Equipment Total"; Tab(30); FormatCurrency(equipmenttotal)
    picResults.Print "Gym Total"; Tab(30); FormatCurrency(gymtotal)
    picResults.Print "********************************************************************"
    picResults.Print "Subtotal"; Tab(30); FormatCurrency(runningtotal)
    tax = runningtotal * 0.07
    grandtotal = runningtotal + tax
    picResults.Print "Tax"; Tab(30); FormatCurrency(tax)
    picResults.Print "Grand Total"; Tab(30); FormatCurrency(grandtotal)

      
End Sub

Private Sub Form_Load()
'This code centers the form on computer screen upon loading.
'this code discovered from Cassie Scherer and Jordan Schmaltz project of developing a vacation

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
