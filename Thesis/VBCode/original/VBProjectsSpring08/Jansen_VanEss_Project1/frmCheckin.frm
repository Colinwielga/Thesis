VERSION 5.00
Begin VB.Form frmCheckin 
   Caption         =   "Check-in Main Menu"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   Picture         =   "frmCheckin.frx":0000
   ScaleHeight     =   9930
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRoomType 
      Height          =   285
      Left            =   3840
      TabIndex        =   25
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox txtNights 
      Height          =   285
      Left            =   3000
      TabIndex        =   23
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox txtHomePhone3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4680
      TabIndex        =   21
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txtHomePhone2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      TabIndex        =   20
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox txtLicensePlate 
      Height          =   285
      Left            =   3000
      TabIndex        =   18
      Top             =   7080
      Width           =   4695
   End
   Begin VB.CommandButton cmdRoomSize 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Finish"
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox txtCar 
      Height          =   285
      Left            =   3000
      TabIndex        =   15
      Top             =   6600
      Width           =   4695
   End
   Begin VB.TextBox txtHomePhone1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox txtAddress3 
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   4920
      Width           =   4695
   End
   Begin VB.TextBox txtAddress2 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   4440
      Width           =   4695
   End
   Begin VB.TextBox txtAddress1 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   3960
      Width           =   4695
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   4440
      X2              =   4440
      Y1              =   8040
      Y2              =   7680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   2160
      X2              =   4440
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   31
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label lblMasterSuite 
      BackStyle       =   0  'Transparent
      Caption         =   "MasterSuite"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label lblSmallSuite 
      BackStyle       =   0  'Transparent
      Caption         =   "SmallSuite"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   840
      TabIndex        =   29
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label lblKing 
      BackStyle       =   0  'Transparent
      Caption         =   "King"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label lblQueen 
      BackStyle       =   0  'Transparent
      Caption         =   "Queen"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label lblDouble 
      BackStyle       =   0  'Transparent
      Caption         =   "DoubleBed"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label lblNightsandRoom 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Nights Stay, Room Type-->"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   24
      Top             =   7560
      Width           =   2865
   End
   Begin VB.Label lblInfoRequest 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "If Purchasing Room, Please Type Personal Information!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   2640
      TabIndex        =   22
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label lblLicensePlate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "License Plate Number -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label lblCar 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Make and Model -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label lblHomePhone 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblOther 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Other Information"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      Top             =   5520
      Width           =   4695
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Label lblAddress3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblAddress2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City, State -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   4440
      UseMnemonic     =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblAddress1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Street Address / Apt # -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblLastName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblFirstName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " First Name -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "frmCheckin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project: Hotel Checkin
'Form: Checkin
'Authors: Ellen Jansen & Stuart Van Ess
'Date: March 28, 2008
'Purpose:   The Main goal of this program is to be used behind the front desk
'           of a hotel in order to check people in and out of the hotel.

'           The purpose of this form is to get information from the user, and
'           Write that information into Arrays in a file.

Option Explicit
Private Sub Label1_Click()

End Sub

Private Sub cmdMain_Click()
    frmCheckin.Hide
    frmMainMenu.Show
End Sub

Private Sub cmdRoomSize_Click()
'Makes the variables = what is put in the text boxes
    FirstName = txtFirstName.Text
    LastName = txtLastName.Text
    Address1 = txtAddress1.Text
    Address2 = txtAddress2.Text
    Address3 = txtAddress3.Text
    HomePhone1 = txtHomePhone1.Text
    HomePhone2 = txtHomePhone2.Text
    HomePhone3 = txtHomePhone3.Text
    Car = txtCar.Text
    LicensePlate = txtLicensePlate.Text
    Room = txtRoomType.Text
    Nights = txtNights.Text
        
'opens our text file to be written on
    Open App.Path & "\Guests.txt" For Append As #2
    
 'writes the information from the text boxes onto the text file named Guests.txt
        Write #2, FirstName, LastName, Address1, Address2, Address3, HomePhone1, HomePhone2, HomePhone3, Car, LicensePlate, Nights, Room
 
 'closes the text file
    Close #2
    
 'Makes all the text boxes blank again, so the next time a person is checking in
 'they do not have the information of the previous customer.
    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtAddress3.Text = ""
    txtHomePhone1.Text = ""
    txtHomePhone2.Text = ""
    txtHomePhone3.Text = ""
    txtCar.Text = ""
    txtLicensePlate.Text = ""
    txtNights.Text = ""
    txtRoomType.Text = ""
    frmMainMenu.Show
    frmCheckin.Hide
End Sub

