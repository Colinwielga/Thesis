VERSION 5.00
Begin VB.Form frmRooms2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Alaskan Rooms"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5415
      Left            =   3480
      ScaleHeight     =   5355
      ScaleWidth      =   6315
      TabIndex        =   5
      Top             =   1200
      Width           =   6375
   End
   Begin VB.CommandButton cmdReturn11 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Alaskan Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton cmdPrices 
      BackColor       =   &H80000014&
      Caption         =   "Display the rooms' prices from greatest to least"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdAlphabeticalOrder 
      BackColor       =   &H80000014&
      Caption         =   "Display rooms in alphabetical order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdRoomsandPrices 
      BackColor       =   &H80000014&
      Caption         =   "Display rooms and prices"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblAls 
      BackColor       =   &H00C0FFC0&
      Caption         =   "  Rooms"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmRooms2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmRooms2
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/15/2009
'Objective: This form includes a list of all of the different rooms available to the user as well as the prices
'of those rooms. There are command buttons that list the rooms in alphabetical order and descending order of price.

Option Explicit
Dim Rooms(1 To 100) As String, Prices(1 To 100000) As Single
Dim CTR As Integer
Dim Pass As Integer, Pos As Integer, Temp As String

Private Sub cmdAlphabeticalOrder_Click()
Dim I As Integer
I = 0

picResults.Cls
picResults.Print "Room"; Tab(35); "Price"
picResults.Print "********************************************************"

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Rooms(Pos) > Rooms(Pos + 1) Then
            Temp = Rooms(Pos)
            Rooms(Pos) = Rooms(Pos + 1)
            Rooms(Pos + 1) = Temp
            
            Temp = Prices(Pos)
            Prices(Pos) = Prices(Pos + 1)
            Prices(Pos + 1) = Temp
        End If
    Next Pos
Next Pass


For I = 1 To CTR
    picResults.Print Rooms(I); Tab(35); FormatCurrency(Prices(I), 2)
Next I
End Sub

Private Sub cmdPrices_Click()
Dim J As Integer
J = 0
picResults.Cls
picResults.Print "Room"; Tab(35); "Price"
picResults.Print "********************************************************"

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Prices(Pos) < Prices(Pos + 1) Then
            Temp = Prices(Pos)
            Prices(Pos) = Prices(Pos + 1)
            Prices(Pos + 1) = Temp
            
            Temp = Rooms(Pos)
            Rooms(Pos) = Rooms(Pos + 1)
            Rooms(Pos + 1) = Temp
        End If
    Next Pos
Next Pass


For J = 1 To CTR
    picResults.Print Rooms(J); Tab(35); FormatCurrency(Prices(J), 2)
Next J

End Sub

Private Sub cmdReturn11_Click()
frmRooms2.Hide
frmAlaskanHome.Show
End Sub

Private Sub cmdRoomsandPrices_Click()

CTR = 0
Open App.Path & "\Rooms2.txt" For Input As #1

picResults.Print "Room"; Tab(35); "Price"
picResults.Print "********************************************************"

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Rooms(CTR), Prices(CTR)
    picResults.Print Rooms(CTR); Tab(35); FormatCurrency(Prices(CTR), 2)
Loop

Close #1

cmdRoomsandPrices.Enabled = False
cmdAlphabeticalOrder.Enabled = True
cmdPrices.Enabled = True
End Sub
