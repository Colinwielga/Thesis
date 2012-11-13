VERSION 5.00
Begin VB.Form frmPricing 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Rooms"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6960
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "List"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   600
      ScaleHeight     =   3555
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "Prices are listed from least expensive to most expensive:"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hotel Check In
'frmPricing
'Shannon Hooley
'10/16/09
'this form allows the guest to see the pricing of rooms from lowest to highest

Private Sub cmdList_Click()
Dim Rooms(1 To 30) As String
Dim RoomNo(1 To 30) As String
Dim Price(1 To 30) As String
Dim CTR As Integer
'pulls info from the list of avaliable rooms
Open App.Path & "\RoomList.txt" For Input As #1
CTR = 0
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Rooms(CTR), RoomNo(CTR), Price(CTR)
Loop
picResults.Print "Room Type"; Tab(25); "Room Number"; Tab(43); "Price per Night"
    picResults.Print "******************************************"
Dim pass As Integer, pos As Integer, J As Long, tempRooms As String, tempRoomNo As Single, tempPrice As String
'switch the numbers and stick them into a temp array
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If Price(pos) > Price(pos + 1) Then
            tempPrice = Price(pos)
            Price(pos) = Price(pos + 1)
            Price(pos + 1) = tempPrice
            tempRoomNo = RoomNo(pos)
            RoomNo(pos) = RoomNo(pos + 1)
            RoomNo(pos + 1) = tempRoomNo
            tempRooms = Rooms(pos)
            Rooms(pos) = Rooms(pos + 1)
            Rooms(pos + 1) = tempRooms
        End If
    Next pos
Next pass

'print the new cities
For J = 1 To CTR
    picResults.Print Rooms(J); Tab(25); RoomNo(J); Tab(43); "$"; Price(J)
Next
End Sub

Private Sub cmdReturn_Click()
'brings the guest back to the layout of the rooms
frmPricing.Hide
frmLayout.Show
End Sub
