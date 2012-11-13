VERSION 5.00
Begin VB.Form frmReserveHotel 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   Picture         =   "frmReserveHotel.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   8520
      TabIndex        =   29
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Back to Home"
      Height          =   495
      Left            =   6480
      TabIndex        =   28
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSubmitReservation 
      Caption         =   "Submit Reservation"
      Height          =   495
      Left            =   2520
      TabIndex        =   27
      Top             =   7560
      Width           =   2415
   End
   Begin VB.OptionButton OptionAmerExpress 
      BackColor       =   &H00C0C0C0&
      Caption         =   "American Express"
      Height          =   255
      Left            =   5880
      TabIndex        =   26
      Top             =   5880
      Width           =   1695
   End
   Begin VB.OptionButton OptionDiscover 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Discover"
      Height          =   195
      Left            =   4320
      TabIndex        =   25
      Top             =   5880
      Width           =   1095
   End
   Begin VB.OptionButton OptionMasterCard 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Master Card"
      Height          =   195
      Left            =   3000
      TabIndex        =   24
      Top             =   5880
      Width           =   1215
   End
   Begin VB.OptionButton OptionVisa 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Visa"
      Height          =   195
      Left            =   2040
      TabIndex        =   23
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox txtCardNum 
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   6360
      Width           =   3255
   End
   Begin VB.TextBox txtExpDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   6960
      Width           =   3255
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox txtPhoneNum 
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   4560
      Width           =   3255
   End
   Begin VB.TextBox txtCounty 
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox txtZIP 
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   3360
      Width           =   3255
   End
   Begin VB.ListBox ListState 
      Height          =   645
      ItemData        =   "frmReserveHotel.frx":2D40C
      Left            =   2160
      List            =   "frmReserveHotel.frx":2D4A6
      TabIndex        =   16
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtMailAddress 
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtLastName 
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtFirstName 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblHotelReservation 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hotel Reservations"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5880
      TabIndex        =   30
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblExpDate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Expiration Date:"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblCardNum 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Credit Card Number:"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label lblCardType 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Credit Card Type:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E-mail Address:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Phone Number:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblCountry 
      BackColor       =   &H00C0C0C0&
      Caption         =   "County:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblZip 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ZIP/Postal Code:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblState 
      BackColor       =   &H00C0C0C0&
      Caption         =   "State:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblCity 
      BackColor       =   &H00C0C0C0&
      Caption         =   "City:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mailing Address:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblLastName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblFirstName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "First Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmReserveHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Project Name: Ideal Greek Island
'Form: frmReserveHotel
'Author: Alie Chandler
'Date Writen: 3/23
'Form Objective: This form is for the user to enter their personal identification information in order to reserve their chosen hotel.

Private Sub cmdHome_Click()
    frmHome1.Show
    frmReserveHotel.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSubmitReservation_Click()
    MsgBox "Your hotel was successfully reserved! You will be contacted shortly with more information. Enjoy Greece :)", , "Reservation Successful"
End Sub
