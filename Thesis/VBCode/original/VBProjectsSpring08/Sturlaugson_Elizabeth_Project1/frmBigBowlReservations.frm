VERSION 5.00
Begin VB.Form frmBigBowlReservations 
   BackColor       =   &H00404080&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to Big Bowl "
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16320
      TabIndex        =   13
      Top             =   2520
      Width           =   1815
   End
   Begin VB.OptionButton optRoseville 
      Caption         =   "Roseville"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   12
      Top             =   8520
      Width           =   1455
   End
   Begin VB.OptionButton optMinnetonka 
      Caption         =   "Minnetonka"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   11
      Top             =   9600
      Width           =   1455
   End
   Begin VB.OptionButton optEdina 
      Caption         =   "Edina"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   10
      Top             =   10680
      Width           =   1455
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   5160
      ScaleHeight     =   3555
      ScaleWidth      =   8715
      TabIndex        =   9
      Top             =   1560
      Width           =   8775
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit Reservation Request"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7920
      TabIndex        =   8
      Top             =   7560
      Width           =   3735
   End
   Begin VB.TextBox txtParty 
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txtDate 
      Height          =   645
      Left            =   3000
      TabIndex        =   3
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtMonth 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtTime 
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblStep3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Step 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      TabIndex        =   18
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label lblStep2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Step 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   17
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label lblStep1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   16
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image img03 
      Height          =   2535
      Left            =   15720
      Picture         =   "frmBigBowlReservations.frx":0000
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   2895
   End
   Begin VB.Image img02 
      Height          =   2535
      Left            =   15720
      Picture         =   "frmBigBowlReservations.frx":0DC4
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Image img01 
      Height          =   2685
      Left            =   15720
      Picture         =   "frmBigBowlReservations.frx":1E8C
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   2865
   End
   Begin VB.Label lblLocation 
      Caption         =   "Please pick a location from the following selection"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   7560
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Reservations for Big Bowl"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   19695
   End
   Begin VB.Label lblHowMany 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Number in party "
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lblDateofMonth 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Month (i.e. June)"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Time (i.e. 3:00pm)     "
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   2655
   End
End
Attribute VB_Name = "frmBigBowlReservations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSCI VB Project: Big Bowl
'frmBigBowlReservations
'Elizabeth K. Sturlaugson
'Due Date: Friday, March 28th, 2008

'This form uses text boxes to gather inforamtion in order to make a reservation for the user


Option Explicit


Private Sub cmdBack_Click()
frmBigBowl.Show
frmBigBowlReservations.Hide

End Sub

Private Sub cmdSubmit_Click()
'declare the variables

Dim Time As String
Dim Month As String
Dim DateofMonth As Integer
Dim NumberinParty As Integer
Dim FirstName As String
Dim LastName As String

'asks for the time, month, date and number of persons in the party


Time = txtTime.Text
Month = txtMonth.Text
DateofMonth = txtDate.Text
NumberinParty = txtParty.Text


'notifies parties of eight will have an added 15% gratuity to their bill

If NumberinParty > 8 Then
MsgBox "Please note that a 15% gratuity will be added to your bill.", , "Notice for Parties Larger than 8"

End If




'enter first and last name to hold reservation
FirstName = InputBox("Please enter your first name", "Reservations")
LastName = InputBox("Please enter your last name", "Reservations")


picResults2.Print "Your reservations are at "; Time; " on "; Month; DateofMonth; " for "; NumberinParty; " persons."
picResults2.Print FirstName; LastName; "your request has been submitted and we look forward to your arrival."




End Sub


