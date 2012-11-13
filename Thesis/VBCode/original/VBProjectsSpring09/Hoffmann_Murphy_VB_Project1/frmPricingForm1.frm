VERSION 5.00
Begin VB.Form frmPricingForm1 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF80FF&
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdActivites 
      BackColor       =   &H00FF80FF&
      Caption         =   "Click to estimate activities cost"
      Height          =   735
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTotalRent 
      BackColor       =   &H00FF80FF&
      Caption         =   "Housing Total"
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtNights 
      Height          =   375
      Left            =   4680
      TabIndex        =   31
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtCabin 
      Height          =   375
      Left            =   2160
      TabIndex        =   30
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   7335
      Left            =   2880
      Picture         =   "frmPricingForm1.frx":0000
      ScaleHeight     =   7275
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   -3480
      Width           =   4815
      Begin VB.Label Label26 
         Caption         =   "Office"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   5520
         Width           =   495
      End
      Begin VB.Shape Shape22 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   1095
         Left            =   120
         Top             =   5880
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "21"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "20"
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "19"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   5760
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "18"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   6000
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "17"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   6120
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "16"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   6480
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "15"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   6720
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "14"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   6960
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "13"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   6600
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "12"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "11"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   6000
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "10"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   5640
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "9"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "8"
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   5160
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   "7"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   5160
         Width           =   135
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   5280
         Width           =   15
      End
      Begin VB.Label Label9 
         Caption         =   "6"
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   5040
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "5"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   5040
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "4"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   4560
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "3"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   4680
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "2"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   4440
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   135
      End
      Begin VB.Shape Shape21 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   1800
         Top             =   6960
         Width           =   135
      End
      Begin VB.Shape Shape20 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   1800
         Top             =   5520
         Width           =   135
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   1800
         Top             =   6480
         Width           =   135
      End
      Begin VB.Shape Shape18 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   1200
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   1200
         Top             =   5280
         Width           =   135
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   1200
         Top             =   5760
         Width           =   135
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   1800
         Top             =   6000
         Width           =   135
      End
      Begin VB.Shape Shape14 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   1200
         Top             =   6720
         Width           =   135
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2880
         Top             =   6360
         Width           =   135
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2880
         Top             =   6720
         Width           =   135
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   3000
         Top             =   5760
         Width           =   135
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2880
         Top             =   6000
         Width           =   135
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   3000
         Top             =   5400
         Width           =   135
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2880
         Top             =   5160
         Width           =   135
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   480
         Top             =   3720
         Width           =   255
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2640
         Top             =   5040
         Width           =   135
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2400
         Top             =   4920
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   480
         Top             =   4320
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   1680
         Top             =   4920
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   1800
         Top             =   4320
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   1200
         Shape           =   1  'Square
         Top             =   4440
         Width           =   255
      End
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Nights"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   29
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Cabin Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   720
      TabIndex        =   28
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Inland Cabins 14-20 $120 a night"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lake Cabins 6-13 $200 a night"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lake Condos 1-5 $300 a night"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblEnter 
      BackStyle       =   0  'Transparent
      Caption         =   "    Enter which      cabin you would     like to stay in,       and how many            nights."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPricingForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Big Sky Resort
'frmPricingForm1
'Ryan Hoffmann and Jamison Murphy
'Written on March 19, 2009
'This Form let's the user decide which area they would like to live
'in and how much it would cost
Option Explicit

'This hides the current form and moves onto the activities form
Private Sub cmdActivites_Click()
    frmActivitiesForm.Show
    frmPricingForm1.Hide
End Sub
'This command goes back to previous form
Private Sub cmdBack_Click()
    frmHomeForm.Show
    frmPricingForm1.Hide
End Sub

Private Sub cmdTotalRent_Click()

'Declaration of variables
Dim CabinSelection As Integer, Nights As Integer
Dim CabinPrice As Single

'Assign textboxes to variables
CabinSelection = txtCabin.Text
Nights = txtNights.Text

'This select case will find out the cost of a night's stay
'at the user's desired location

Select Case CabinSelection
    Case Is < 1
        MsgBox "Please Select a numbered cabin we have", , "Error"
    Case Is < 6
        CabinPrice = 300
    Case Is < 14
        CabinPrice = 200
    Case Is < 22
        CabinPrice = 120
    Case Else
        MsgBox "Please Select a numbered cabin we have", , "Error"
End Select

'This will find the total based on how many nights the user
'wants to stay
LodgingTotal = CabinPrice * Nights
        
'Prints a message box telling user the price
MsgBox "Your predicted lodging total is " & FormatCurrency(LodgingTotal) & ".", , "Lodging Total"
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub

