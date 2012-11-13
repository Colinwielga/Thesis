VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   Picture         =   "Info.frx":0000
   ScaleHeight     =   7395
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumNights 
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdFinish 
      BackColor       =   &H00808080&
      Caption         =   "Finish Your Check In"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox txtRoomChoice 
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox txtLastNumbers 
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtFirstNumbers 
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtAreaCode 
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtZipCode 
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtState 
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtLastName 
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtFirstName 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Check In"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label lblNights 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of nights:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblRoom 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Choice:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblPhone 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lblStateZip 
      BackStyle       =   0  'Transparent
      Caption         =   "State, Zip Code:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblCity 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Address:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblLast 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblFirst 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Let's grab some information from you:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hotel Check In
'frmInfo
'Shannon Hooley
'10/16/09
'This form lets the guest enter their info for a more personalized bill

Private Sub cmdFinish_Click()
Dim FirstName As String
Dim LastName As String
Dim Address As String
Dim City As String
Dim State As String
Dim ZipCode As String
Dim AreaCode As String
Dim FirstNumbers As String
Dim LastNumbers As String
Dim RoomChoice As String
Dim NumNights As String
'tells the computer what the text boxes are for
FirstName = txtFirstName.Text
LastName = txtLastName.Text
Address = txtAddress.Text
City = txtCity.Text
State = txtState.Text
ZipCode = txtZipCode.Text
AreaCode = txtAreaCode.Text
FirstNumbers = txtFirstNumbers.Text
LastNumbers = txtLastNumbers.Text
RoomChoice = txtRoomChoice.Text
NumNights = txtNumNights.Text

'writes the info from the text boxes into a notepad
Open App.Path & "\Info.txt" For Append As #2
    Write #2, FirstName, LastName, Address, City, State, ZipCode, AreaCode, FirstNumbers, LastNumbers, RoomChoice, NumNights
Close #2

'clears the info from the past guest
txtFirstName.Text = ""
txtLastName.Text = ""
txtAddress.Text = ""
txtCity.Text = ""
txtState.Text = ""
txtZipCode.Text = ""
txtAreaCode.Text = ""
txtFirstNumbers.Text = ""
txtLastNumbers.Text = ""
txtRoomChoice.Text = ""
txtNumNights.Text = ""
frmInfo.Hide
frmRoomNumber.Show
End Sub

Private Sub cmdReturn_Click()
'brings the guest back to the check in area
frmInfo.Hide
frmCheckIn.Show
End Sub
