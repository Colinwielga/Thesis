VERSION 5.00
Begin VB.Form frmCheckOut 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "You aren't ready to Check Out"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add up your bill"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   600
      ScaleHeight     =   7995
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   1080
      Width           =   6255
   End
   Begin VB.CommandButton cmdLeave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leave the Lake Side Inn"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   3255
   End
   Begin VB.Label lblBill 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Out from the Lake Side Inn"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hotel Check In
'frmCheckOut
'Shannon Hooley
'10/16/09
'This is where the guest can check out and see their bill from their stay

Private Sub cmdLeave_Click()
'this ends the prgram
End
End Sub

Private Sub cmdReturn_Click()
'brings the guest back to their own room description
frmCheckOut.Hide
frmRoomNumber.Show
End Sub

Private Sub cmdTotal_Click()
'dims the array from the info the guest entered in
    Dim TotalCTR As Integer
    Dim FirstNameArray(1 To 100) As String
    Dim LastNameArray(1 To 100) As String
    Dim AddressArray(1 To 100) As String
    Dim CityArray(1 To 100) As String
    Dim StateArray(1 To 100) As String
    Dim ZipCodeArray(1 To 100) As String
    Dim AreaCodeArray(1 To 100) As String
    Dim FirstNumbersArray(1 To 100) As String
    Dim LastNumbersArray(1 To 100) As String
    Dim RoomChoiceArray(1 To 100) As String
    Dim NumNightsArray(1 To 100) As String
    Dim Bill As Single
    Dim Tax As Single
    Dim Total As Single
    Dim FirstNameBox As String
    Dim LastNameBox As String
    Dim Found As Boolean
    Dim pos As Integer
'says that you haven't compared any info yet
Found = False
'clears the picture box
picResults.Cls
'gets the guest's info from them
FirstNameBox = InputBox("What is your first name?", "First Name", "")
LastNameBox = InputBox("What is your last name?", "LastName", "")
'grabs data from txt file
Open App.Path & "\Info.txt" For Input As #2
    TotalCTR = 0
Do While Not EOF(2)
    TotalCTR = TotalCTR + 1
    Input #2, FirstNameArray(TotalCTR), LastNameArray(TotalCTR), AddressArray(TotalCTR), CityArray(TotalCTR), StateArray(TotalCTR), ZipCodeArray(TotalCTR), AreaCodeArray(TotalCTR), FirstNumbersArray(TotalCTR), LastNumbersArray(TotalCTR), RoomChoiceArray(TotalCTR), NumNightsArray(TotalCTR)
Loop
Close #2
'compares the info entered in by the guest to the data they entered in previously
Do While (Not Found) And (pos < TotalCTR)
        pos = pos + 1
        If FirstNameArray(pos) = FirstNameBox And LastNameArray(pos) = LastNameBox Then
            Found = True
'prints the bill
picResults.Print "Bill for: "
picResults.Print "*************************"
picResults.Print FirstNameArray(pos); Tab(10); LastNameArray(pos)
picResults.Print
picResults.Print AddressArray(pos)
picResults.Print CityArray(pos); Tab(20); ","; StateArray(pos); Tab(27); ZipCodeArray(pos)
picResults.Print
picResults.Print AreaCodeArray(pos); "-"; FirstNumbersArray(pos); "-"; LastNumbersArray(pos)
picResults.Print
picResults.Print RoomChoiceArray(pos)
picResults.Print
'puts a price to the roomchoice
If RoomChoiceArray(pos) = "Presidential Suite" Then
    RoomChoiceArray(pos) = "$453"
ElseIf RoomChoiceArray(pos) = "Suite" Then
    RoomChoiceArray(pos) = "$379"
ElseIf RoomChoiceArray(pos) = "King" Then
    RoomChoiceArray(pos) = "$206"
ElseIf RoomChoiceArray(pos) = "Queen" Then
    RoomChoiceArray(pos) = "$150"
ElseIf RoomChoiceArray(pos) = "Double" Then
    RoomChoiceArray(pos) = "$109"
End If
'adds up the total using the info entered into the array
Bill = NumNightsArray(pos) * RoomChoiceArray(pos)
Tax = Bill * 0.07
Total = Bill + Tax
'print the total
picResults.Print "Your subtotal is "; FormatCurrency(Bill); "."
picResults.Print
picResults.Print "Tax is "; FormatCurrency(Tax); "."
picResults.Print
picResults.Print "Your total is "; FormatCurrency(Total); "."

    End If
Loop
End Sub
