VERSION 5.00
Begin VB.Form frmHotel 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   17625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd2 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
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
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdNext2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Now Rent A Car For Cheap!"
      Height          =   735
      Left            =   14880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7560
      Width           =   2655
   End
   Begin VB.PictureBox picTotal 
      Height          =   2055
      Left            =   6600
      ScaleHeight     =   1995
      ScaleWidth      =   6075
      TabIndex        =   12
      Top             =   6240
      Width           =   6135
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirm Your Reservation!"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   3495
   End
   Begin VB.CommandButton cmdReserve 
      BackColor       =   &H000000FF&
      Caption         =   "Reserve Room"
      Height          =   975
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   2175
      Left            =   6000
      ScaleHeight     =   2115
      ScaleWidth      =   11115
      TabIndex        =   9
      Top             =   3960
      Width           =   11175
   End
   Begin VB.TextBox txtRoomNumber 
      Height          =   1335
      Left            =   2880
      TabIndex        =   8
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox txtNights 
      Height          =   975
      Left            =   2880
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox txtGuest 
      Height          =   1095
      Left            =   2880
      TabIndex        =   5
      Top             =   3000
      Width           =   2535
   End
   Begin VB.PictureBox picResult 
      Height          =   2055
      Left            =   6000
      ScaleHeight     =   1995
      ScaleWidth      =   10995
      TabIndex        =   4
      Top             =   1680
      Width           =   11055
   End
   Begin VB.CommandButton cmdRooms 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Click Here For Room Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "The Room Number You Would Like"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label lblNights 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Number of Nights Staying"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Number of Guest In The Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblPricing 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Lets Talk Money"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "frmHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hotel Page'
'For this page i referenced the VB example project done by cassie and jordan'
'they helped me with structure,
'all patterns,the numbers, and ordering were doner by me'
'this is the most complicated frm of my project
'expect to see arrays, if then, case statements, file inputs/outputs, and much more'
'october 14th-18th'
'Blake Bauer'

Option Explicit
Dim Nights As Integer, RoomNumber As Integer
Dim Total As Single, Tax As Single
Dim Ctr As Integer

'quit button'
Private Sub cmdEnd2_Click()
    End
End Sub

'hiding and showing forms'
Private Sub cmdNext2_Click()
    
    
    frmHotel.Hide
    frmCars.Show
    

    
End Sub

Private Sub cmdReserve_Click()
Dim Guests As Integer

picResults.Cls

'decalring variables'


'creating textbox and what they are'

Guests = txtGuest.Text
Nights = txtNights.Text
RoomNumber = txtRoomNumber.Text

'making my case if statements'
'this is info from the textbox'

If RoomNumber = 1 Then
        Total = Nights * 150
    ElseIf RoomNumber = 2 Then
        Total = Nights * 150
    ElseIf RoomNumber = 3 Then
        Total = Nights * 200
    ElseIf RoomNumber = 4 Then
        Total = Nights * 250
    ElseIf RoomNumber = 5 Then
        Total = Nights * 1000
End If

'making the price suspetiable to taxing'
Tax = 0.1 * Total
TabTotal = Total + Tax


'Very large If then statements that create a visual aid that helps you decide what room and price'
'making many different ranges possible'

If RoomNumber = 1 And Guests <= 6 Then
    picResults.Print "You Have reserved the two twins bed room for "; Guests; " people for "; Nights; " nights. Your Hotel cost will be "; FormatCurrency(TabTotal); " Tax Included."
                cmdYes.Visible = True
End If

If RoomNumber = 2 And Guests <= 4 Then
    picResults.Print "You Have reserved the One Queen Bed Room for "; Guests; " people for "; Nights; " nights. Your Hotel cost will be "; FormatCurrency(TabTotal); ". "
          cmdYes.Visible = True
End If
    
If RoomNumber = 3 And Guests <= 4 Then

 picResults.Print "You Have reserved the One King Bed Room for "; Guests; " people for "; Nights; " nights. Your Hotel cost will be "; FormatCurrency(TabTotal); "."
          cmdYes.Visible = True
End If

If RoomNumber = 4 And Guests <= 5 Then
     picResults.Print "You Have reserved the Suite Delux Room for "; Guests; " people for "; Nights; " nights. Your Hotel cost will be "; FormatCurrency(TabTotal); "."
          cmdYes.Visible = True
End If

If RoomNumber = 5 And Guests <= 20 Then
     picResults.Print "You Have reserved the House! for "; Guests; " people for "; Nights; " nights. Your Hotel cost will be "; FormatCurrency(TabTotal); "."
          cmdYes.Visible = True
End If

If RoomNumber = 1 And Guests > 6 Then
     MsgBox ("Sorry your number of guests exceeds the capacity of your chosen room type.")
     
End If

If RoomNumber = 2 And Guests > 4 Then
    MsgBox ("Sorry your number of guests exceeds the capacity of your chosen room type.")
     
End If

If RoomNumber = 3 And Guests > 4 Then
    MsgBox ("Sorry your number of guests exceeds the capacity of your chosen room type.")
     
End If

If RoomNumber = 4 And Guests > 5 Then
    MsgBox ("Sorry your number of guests exceeds the capacity of your chosen room type.")
     
End If

If RoomNumber = 5 And Guests > 20 Then
    MsgBox ("Sorry your number of guests exceeds the capacity of your chosen room type.")
     
End If


End Sub



Private Sub cmdRooms_Click()
'declaring variabls for my printing in the pic box'
Dim roomtype(1 To 10) As String
Dim Price(1 To 100) As Single
Dim capacity(1 To 10) As Integer
Dim RoomNumber(1 To 10) As Integer
Dim J As Integer

'opening my text file'
Open App.Path & "\Rooms.txt" For Input As #1


Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, RoomNumber(Ctr)
    Input #1, roomtype(Ctr)
    Input #1, Price(Ctr)
    Input #1, capacity(Ctr)
    
Loop
Close #1

picResult.Print "Room Number"; , , ; "Room Type"; , , "Price Per Night"; , , ; "Room Capacity"
picResult.Print "*****************************************************************************************************************************************************************************************"

'printing my file'
For J = 1 To Ctr
    picResult.Print RoomNumber(J); , , ; "   ", roomtype(J); , , ; FormatCurrency(Price(J)); , , ; , , ; capacity(J)
Next J
'making cmd buttons go away'
cmdReserve.Visible = True
cmdRooms.Visible = False
cmdYes.Visible = False
cmdNext2.Visible = False



End Sub

Private Sub cmdYes_Click()

picTotal.Print "Your reservation has been confirmed!"
picTotal.Print "_______________________________________"


'text boxes'
RoomNumber = txtRoomNumber.Text
Nights = txtNights.Text


'My case statements'
'doing math for the nights being there'
Select Case RoomNumber
    Case Is = 1
        Total = Nights * 150
    Case Is = 2
        Total = Nights * 150
    Case Is = 3
        Total = Nights * 200
    Case Is = 4
        Total = Nights * 250
    Case Is = 5
        Total = Nights * 1000
End Select



'creating a room tax'
Tax = 0.1 * Total
TabTotal = Total + Tax


'printong results'
picTotal.Print "SubTotal:", FormatCurrency(Total)
picTotal.Print "Tax:", FormatCurrency(Tax)
picTotal.Print "*********************************************"
picTotal.Print "Total:", FormatCurrency(TabTotal)
'changing frm's'
cmdYes.Visible = False
cmdNext2.Visible = True



End Sub

