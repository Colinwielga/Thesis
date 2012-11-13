VERSION 5.00
Begin VB.Form frmShop 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hockey Lodge"
   ClientHeight    =   11265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   ScaleHeight     =   11265
   ScaleWidth      =   14325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00004000&
      Caption         =   "Back To The Rink!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10080
      Width           =   4575
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00000080&
      Caption         =   "Clear The List"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9120
      Width           =   4575
   End
   Begin VB.CommandButton cmdRunningTotal 
      BackColor       =   &H0057C0E8&
      Caption         =   "Calculate The Total Cost"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   4575
   End
   Begin VB.CommandButton cmdPuck 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   6000
      Picture         =   "frmShop.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   1695
      Left            =   4080
      Picture         =   "frmShop.frx":0CAA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   1695
      Left            =   2160
      Picture         =   "frmShop.frx":1C5B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdHat 
      Height          =   1695
      Left            =   240
      Picture         =   "frmShop.frx":3007
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0057C0E8&
      Height          =   7215
      Left            =   9480
      ScaleHeight     =   7155
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label lblPrice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   12000
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblProduct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblClick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on the product you want to purchase:"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   5535
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome To The Hockey Lodge! "
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Image imgLodge 
      Height          =   2415
      Left            =   5160
      Picture         =   "frmShop.frx":3D2A
      Top             =   360
      Width           =   4200
   End
   Begin VB.Image imgTable 
      Height          =   4590
      Left            =   720
      Picture         =   "frmShop.frx":62D5
      Top             =   6000
      Width           =   6120
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Wild Visual Basic project
'Hockey Lodge Form
'Authors: Adam Andersson and Patrick Johnson
'24 Feb 2010
'The purpose of this form is so the user can shop for
'Minnesota Wild souvenirs and see how much they cost


Dim RunningTotal As Single, Total As Single

'this button adds a Minnesota Wild hat to the purchase
Private Sub cmdHat_Click()
Dim Hat As Single
Hat = 12.99
picResults.Print "Minnesota Wild Hat"; Tab(40); FormatCurrency(Hat)
RunningTotal = RunningTotal + Hat
End Sub

'this button adds a Minnesota Wild jersey to the purchase
Private Sub cmdJersey_Click()
Dim Jersey As Single
Jersey = 65.99
picResults.Print "Green Minnesota Wild Jersey"; Tab(40); FormatCurrency(Jersey)
RunningTotal = RunningTotal + Jersey
End Sub

'this button adds a Minnesota Wild puck to the purchase
Private Sub cmdPuck_Click()
Dim Puck As Single
Puck = 9.99
picResults.Print "Wild Puck"; Tab(40); FormatCurrency(Puck)
RunningTotal = RunningTotal + Puck
End Sub

'this button adds a Minnesota wild shirt to the purchase
Private Sub cmdShirt_Click()
Dim Shirt As Single
Shirt = 19.99
picResults.Print "Red Minnesota Wild Shirt"; Tab(40); FormatCurrency(Shirt)
RunningTotal = RunningTotal + Shirt
End Sub

'this button calculates the running total, adds a 7% tax to the purchase,
'then calculates the total and displays it in a picture box
Private Sub cmdRunningTotal_Click()
Dim Tax As Single
Tax = RunningTotal * 0.07
Total = RunningTotal + Tax

'print buffer zone
picResults.Print
picResults.Print "**************************************************************"
picResults.Print FormatCurrency(Total)

'begin select case statement
Select Case Total
    Case Is >= 100
        picResults.Print "Wow, that is soooo expensive!"
    Case Is >= 75
        picResults.Print "Holy macaroni, you are really a wild fan bud!"
    Case Is >= 50
        picResults.Print "You must really like the Minnesota Wild."
    Case Is >= 30
        picResults.Print "Hmmmm, come back and buy some more later..."
    Case Else
        picResults.Print "You are a cheap..."
End Select
End Sub

'this button clears the picture box
Private Sub cmdClear_Click()
picResults.Cls
RunningTotal = 0
Total = 0
End Sub

'this button takes the user back to the Main Form
Private Sub cmdBack_Click()
'show the main form and hide the other forms
frmWelcome.Hide
frmMain.Show
frmRoster.Hide
frmShot.Hide
frmLeague.Hide
frmShop.Hide
End Sub
