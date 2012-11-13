VERSION 5.00
Begin VB.Form frmSorting 
   Caption         =   "Sorting Form"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   Picture         =   "frmSorting.frx":0000
   ScaleHeight     =   7890
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Price You Are Going To Pay From The Table Above"
      Height          =   1455
      Left            =   1560
      TabIndex        =   4
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdListPriceOptions 
      Caption         =   "List Room Options and Prices Starting From the Least Expensive"
      Height          =   1455
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdAvailable 
      Caption         =   "Find Features of Desired Room and Availability"
      Height          =   1455
      Left            =   1560
      TabIndex        =   2
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton cmdCheapest 
      Caption         =   "Show the Cheapest Room"
      Height          =   1455
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.PictureBox picSort 
      Height          =   6495
      Left            =   4920
      ScaleHeight     =   6435
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
End
Attribute VB_Name = "frmsorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Hotel Checkin
'Form: Sorting
'Authors: Ellen Jansen & Stuart Van Ess
'Date: March 28, 2008
'Purpose:   On this page, we are able to sort the hotel rooms many different
'           ways. We can simply show the cheapest, show the different price
'           options, or input how much they are paying and print the room
'           they are staying in.
    



Option Explicit
Dim CTR As Integer
Private Sub cmdAvailable_Click()

'shows the room size form and hides the sorting form
    frmsorting.Hide
    frmRoomSize.Show
End Sub

Private Sub cmdCheapest_Click()
    
    Dim I As Integer
    Dim PlaceValue As Single
    Dim PlaceRoom As String
   
'Sets the room name as blank, and the place value as a large value which will
'be replaced with the smaller ones in the arrays
    PlaceRoom = ""
    PlaceValue = 999
    
 'starts the"I" counter at one
    I = 1
    
'For-Next loop that searches the extent of the list to find the lowest value
'for price.
    For I = 1 To 5
        If Price(I) < PlaceValue Then
        PlaceValue = Price(I)
        End If
    Next I
    
'clears the printing screen and
'prints the lowest value
    picSort.Cls
    picSort.Print "The cheapest room option is a double which costs "; PlaceValue; "."
    
End Sub
Private Sub cmdListPriceOptions_Click()
 
    Dim Pass As Single
    Dim Pos As Single
    Dim TempPrice As Single
    Dim TempRoom As String
    Dim K As Integer


'sets up the array from lowest price to highest, and keeps the price and room
'name consistent with each other
    For Pass = 1 To 5 - 1
        For Pos = 1 To 5 - Pass
            If Price(Pos) > Price(Pos + 1) Then
                TempPrice = Price(Pos)
                Price(Pos) = Price(Pos + 1)
                Price(Pos + 1) = TempPrice
                TempRoom = Rooms(Pos)
                Rooms(Pos) = Rooms(Pos + 1)
                Rooms(Pos + 1) = TempRoom
            End If
        Next Pos
    Next Pass
    
'clears the printing screen and
'prints the headings
    picSort.Cls
    picSort.Print "Rooms", "  Prices"
    picSort.Print "------------------------------------------------------"
    
'prints the array in a table format
    For K = 1 To 5
        picSort.Print Rooms(K), Price(K)
    Next K

End Sub
Private Sub cmdSelect_Click()

Dim PriceRoom As Single

'asks the user to input how much they are paying for their room
PriceRoom = InputBox("Please enter a price from the table of available prices.")

picSort.Print ""


'uses the select case format to determine which room the person is staying in,
'according to their price entered.
Select Case PriceRoom
    Case Is = 119.99
        picSort.Print "You're getting a double!"
    Case Is = 139.99
        picSort.Print "You're getting a queen!"
    Case Is = 149.99
        picSort.Print "You're getting a king!"
    Case Is = 179.99
        picSort.Print "You're getting a small-suite!"
    Case Is = 199.99
        picSort.Print "You're getting a master-suite!"
    Case Else
        picSort.Print "Sorry. You entered an unacceptable price."
End Select



End Sub

Private Sub Form_Load()


'when the form initially loads, the text file named roomsandrates gets
'read into 2 separate arrays.
Open App.Path & "\roomsandrates.txt" For Input As #1
        
        CTR = 1
        
        Do While Not EOF(1)
            Input #1, Rooms(CTR), Price(CTR)
            CTR = CTR + 1
        Loop
Close #1

End Sub
