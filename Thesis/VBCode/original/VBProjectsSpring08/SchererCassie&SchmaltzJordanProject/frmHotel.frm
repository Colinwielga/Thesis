VERSION 5.00
Begin VB.Form frmHotel 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   6840
      ScaleHeight     =   1635
      ScaleWidth      =   3915
      TabIndex        =   15
      Top             =   7440
      Width           =   3975
   End
   Begin VB.CommandButton cmdload 
      BackColor       =   &H00FF00FF&
      Caption         =   "Load Room Options"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdactivities 
      BackColor       =   &H00FF00FF&
      Caption         =   "Continue to Next Page to Plan Activities!"
      Height          =   1335
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdconfirm 
      BackColor       =   &H00FF8080&
      Caption         =   "Confirm Reservation"
      Height          =   855
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdbookIt 
      BackColor       =   &H00FF8080&
      Caption         =   "Book Room"
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox txtroom 
      Height          =   855
      Left            =   3480
      TabIndex        =   5
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtnights 
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtguests 
      Height          =   855
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   6480
      ScaleHeight     =   2115
      ScaleWidth      =   8355
      TabIndex        =   2
      Top             =   4440
      Width           =   8415
   End
   Begin VB.PictureBox picTable 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   6480
      ScaleHeight     =   2235
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   1560
      Width           =   8415
   End
   Begin VB.Label lblload 
      BackColor       =   &H00FF00FF&
      Caption         =   "<= Press Here First to Start!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label lblreceipt 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Your Receipt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label lbltable 
      BackColor       =   &H00C0C0FF&
      Caption         =   "These are the options you have for room selection."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lblroom 
      BackColor       =   &H00000000&
      Caption         =   "Please enter the corresponding number to the type of room you would like."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label lblnights 
      BackColor       =   &H00000000&
      Caption         =   "Please enter the number of nights you will be staying."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label lblguests 
      BackColor       =   &H00000000&
      Caption         =   "Please enter the number of guests in the room."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label lbldirections 
      BackColor       =   &H000000FF&
      Caption         =   "If you are booking multiply rooms, please book each room separately."
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "frmHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmHotel
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/6/08 - 3/7/08
'Objective: The user makes a hotel reservation.
'They input how many guests, what type of room they want, and for how many nights
'If they enter a number of guests that exceeds the room capacity the program will tell them to reenter a number of guests
'We inputed a file and then formatted it into an array
'We used input boxes, message boxes, and if then statements


Option Explicit
Dim nights As Integer
Dim roomtype As Integer
Dim subtotal As Single
Dim tax As Single

Private Sub cmdactivities_Click()

'Here the user goes to another form to select the activities they wish to do on vacation

frmHotel.Hide
frmActivites.Show

End Sub


Private Sub cmdbookIt_Click()

'Here we declared our variables

Dim guests As Integer

'Here we assigned our variables the numbers that are input from the user in the form of text boxes

guests = txtguests.Text
nights = txtnights.Text
roomtype = txtroom.Text

'Here we used an if then else statement to determine the subtotal of the particular room they have chosen

If roomtype = 1 Then
        subtotal = nights * 120
    ElseIf roomtype = 2 Then
        subtotal = nights * 157
    ElseIf roomtype = 3 Then
        subtotal = nights * 145
    ElseIf roomtype = 4 Then
        subtotal = nights * 215
End If

'Using the subtotal we calculated from the previous if then else statement we derived the total cost of the room

tax = 0.2 * subtotal
Hoteltotal = subtotal + tax

'This monster if then else statement determines if their number of guests can fit in their chosen room type
'Each room has a maximum capacity and this code makes sure that the individual party does not excede this capacity
'If the party is within the capacity it prints the type of room and number of people staying in the room along with the cost and number of nights stayed
'Only if the user books a valid room will the confirm registration button appear

If roomtype <= 2 And guests <= 2 Then
        If roomtype = 1 Then
                picResults.Print "You have booked a Basic room for "; guests; " people for "; nights; " nights. The total cost of your stay is "; FormatCurrency(Hoteltotal); "."
                cmdconfirm.Visible = True
            ElseIf roomtype = 2 Then
                picResults.Print "You have booked a King room for "; guests; " people for "; nights; " nights. The total cost of your stay is "; FormatCurrency(Hoteltotal); "."
                cmdconfirm.Visible = True
        End If
    ElseIf roomtype <= 2 And guests > 2 Then
        MsgBox ("Sorry your number of guests exceeds the capacity of your chosen room type.")
    ElseIf roomtype <= 4 And guests <= 4 Then
        If roomtype = 1 Then
                picResults.Print "You have booked a Basic room for "; guests; " people for "; nights; " nights. The total cost of your stay is "; FormatCurrency(Hoteltotal); "."
                cmdconfirm.Visible = True
            ElseIf roomtype = 2 Then
                picResults.Print "You have booked a King room for "; guests; " people for "; nights; " nights. The total cost of your stay is "; FormatCurrency(Hoteltotal); "."
                cmdconfirm.Visible = True
            ElseIf roomtype = 3 Then
                picResults.Print "You have booked a Standard room for "; guests; " people for "; nights; " nights. The total cost of your stay is "; FormatCurrency(Hoteltotal); "."
                cmdconfirm.Visible = True
            ElseIf roomtype = 4 Then
                picResults.Print "You have booked a Suite for "; guests; " people for "; nights; " nights. The total cost of your stay is "; FormatCurrency(Hoteltotal); "."
                cmdconfirm.Visible = True
        End If
    ElseIf roomtype <= 4 And guests > 4 Then
        MsgBox ("Sorry your number of guests exceeds the capacity of your chosen room type.")
    ElseIf roomtype > 4 Then
        MsgBox ("Sorry you entered an invalid room type number.")
End If


End Sub

Private Sub cmdconfirm_Click()

'Here we are confirming the guests reservation

picResults.Print ""
picResults.Print "Your reservation has been confirmed!"

'Here we declared our variables

roomtype = txtroom.Text
nights = txtnights.Text

'Here we used an if then else statement to determine the subtotal of the particular room they have chosen

Select Case roomtype
    Case Is = 1
        subtotal = nights * 120
    Case Is = 2
        subtotal = nights * 157
    Case Is = 3
        subtotal = nights * 145
    Case Is = 4
        subtotal = nights * 215
End Select


'Using the subtotal we calculated from the previous if then else statement we derived the total cost of the room

tax = 0.2 * subtotal
Hoteltotal = subtotal + tax

'Here we are printing a receipt for the customer

picResults2.Print "Subtotal:", FormatCurrency(subtotal)
picResults2.Print "Tax:", FormatCurrency(tax)
picResults2.Print "*********************************************"
picResults2.Print "Total:", FormatCurrency(Hoteltotal)

'Only after the user confirms their reservation can they look at activites

cmdactivities.Visible = True

End Sub

Private Sub cmdload_Click()

'Here we declared our variables.

Dim number(1 To 5) As Integer
Dim roomtype(1 To 5) As String
Dim price(1 To 5) As Single
Dim capacity(1 To 5) As Integer
Dim description(1 To 5) As String

'Here we opened up a path to the file RoomType

Open App.Path & "\RoomType.txt" For Input As #1

'Here we loaded the file RoomType into an array

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, number(CTR)
    Input #1, roomtype(CTR)
    Input #1, price(CTR)
    Input #1, capacity(CTR)
    Input #1, description(CTR)
Loop

'Here we displayed our array in a table format for the user

picTable.Print "Number", "Room Type", "Price Per Night", "Room Capacity", "Description"

For J = 1 To CTR
    picTable.Print number(J), roomtype(J), FormatCurrency(price(J)), , capacity(J), , description(J)
Next J

'The user can not book their hotel stay before loading this data

cmdbookIt.Visible = True
cmdload.Visible = False
lblload.Visible = False

End Sub

Private Sub cmdquit_Click()
End
End Sub


Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
