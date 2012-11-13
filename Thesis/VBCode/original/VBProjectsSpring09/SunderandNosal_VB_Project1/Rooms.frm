VERSION 5.00
Begin VB.Form frmRooms 
   BackColor       =   &H00808080&
   Caption         =   "Rooms"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBackToHome 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to the Caribbean Home Page"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton CmdPricesInOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Display the room prices from greatest to least "
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
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton cmdAlphabetical 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Display the rooms in alphabetical order"
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
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton cmdRooms 
      BackColor       =   &H00E0E0E0&
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
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      Height          =   6615
      Left            =   4320
      ScaleHeight     =   6555
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label lblRooms 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Rooms"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmRooms
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/15/2009
'Objective: This form includes a list of all of the different rooms available to the user as well as the prices
'of those rooms. There are command buttons that list the rooms in alphabetical order and descending order of price.

Option Explicit
Dim Rooms(1 To 100) As String, Prices(1 To 100000) As Single
Dim CTR As Integer
Dim Pass As Integer, Pos As Integer, Temp As String

Private Sub cmdAlphabetical_Click()
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

Private Sub cmdGoBacktoHome_Click()
frmRooms.Hide
frmCaribbeanHome.Show
End Sub

Private Sub CmdPricesInOrder_Click()
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

Private Sub cmdRooms_Click()

CTR = 0
Open App.Path & "\Rooms.txt" For Input As #1

picResults.Print "Room"; Tab(35); "Price"
picResults.Print "********************************************************"

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Rooms(CTR), Prices(CTR)
    picResults.Print Rooms(CTR); Tab(35); FormatCurrency(Prices(CTR), 2)
Loop

Close #1
cmdRooms.Enabled = False
cmdAlphabetical.Enabled = True
CmdPricesInOrder.Enabled = True
End Sub
