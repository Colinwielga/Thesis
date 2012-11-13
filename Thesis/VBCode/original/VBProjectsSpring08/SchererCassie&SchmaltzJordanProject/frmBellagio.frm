VERSION 5.00
Begin VB.Form frmBellagio 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   Picture         =   "frmBellagio.frx":0000
   ScaleHeight     =   9375
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdload 
      Caption         =   "Load"
      Height          =   855
      Left            =   7560
      TabIndex        =   15
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdactivities 
      Caption         =   "Continue to Next Page to Plan Activities!"
      Height          =   1335
      Left            =   3360
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdconfirm 
      Caption         =   "Confirm Reservation"
      Height          =   855
      Left            =   3360
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdbookIt 
      Caption         =   "Book Room"
      Height          =   855
      Left            =   480
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   3960
      TabIndex        =   10
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox txtroom 
      Height          =   855
      Left            =   3600
      TabIndex        =   5
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txtnights 
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtguests 
      Height          =   855
      Left            =   3600
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      Height          =   4335
      Left            =   6480
      ScaleHeight     =   4275
      ScaleWidth      =   8115
      TabIndex        =   2
      Top             =   4440
      Width           =   8175
   End
   Begin VB.PictureBox picTable 
      Height          =   2295
      Left            =   6720
      ScaleHeight     =   2235
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   1560
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   600
      Picture         =   "frmBellagio.frx":0FBF
      Top             =   6600
      Width           =   2190
   End
   Begin VB.Label lbltable 
      Caption         =   "These are the options you have for room selection."
      Height          =   375
      Left            =   10080
      TabIndex        =   14
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label lblroom 
      Caption         =   "Please enter the corresponding number to the type of room you would like."
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label lblnights 
      Caption         =   "Please enter the number of nights you will be staying."
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label lblguests 
      Caption         =   "Please enter the number of guests in the room."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label lbldirections 
      Caption         =   "if you are booking multiply rooms, please book each room separately."
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "frmBellagio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdactivities_Click()

'Here the user goes to another form to select the activities they wish to do on vacation

frmBellagio.Hide
frmVegasActivities.Show

End Sub

Private Sub cmdback_Click()

'Here we allow the user the option to return to the previous screen

frmBellagio.Hide
frmVegasHotels.Show

End Sub

Private Sub cmdbookIt_Click()

'Here we declared our variables

Dim guests As Integer
Dim nights As Integer
Dim roomtype As Integer

'Here we assigned our variables the numbers that are input from the user in the form of text boxes

guests = txtguests.Text
nights = txtnight.Text
roomtype = txtroomtype.Text



'The user can confirm their reservation only after they have booked it

cmdconfirm.Visible = True

End Sub

Private Sub cmdload_Click()

'Here we declared our variables.

Dim number(1 To 5) As Integer
Dim roomtype(1 To 5) As String
Dim price(1 To 5) As Single
Dim description(1 To 5) As String
Dim CTR As Integer
Dim J As Integer

'Here we opened up a path to the file RoomType

Open App.Path & "\RoomType.txt" For Input As #1

'Here we loaded the file RoomType into an array

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, number(CTR)
    Input #1, roomtype(CTR)
    Input #1, price(CTR)
    Input #1, description(CTR)
Loop

'Here we displayed our array in a table format for the user

picTable.Print "Number", "Room Type", "Price", "Description"

For J = 1 To CTR
    picTable.Print number(J), roomtype(J), price(J), description(J)
Next J

'The user can not book their hotel stay before loading this data

cmdbookIt.Visible = True

End Sub

Private Sub cmdquit_Click()
End
End Sub
