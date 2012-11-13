VERSION 5.00
Begin VB.Form AircraftPrice 
   BackColor       =   &H00FF8080&
   Caption         =   "Aircraft Pricing"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form3"
   ScaleHeight     =   7650
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmainmenu 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8160
      TabIndex        =   4
      Top             =   5880
      Width           =   2295
   End
   Begin VB.PictureBox pbxresults2 
      Height          =   4575
      Left            =   1920
      ScaleHeight     =   4515
      ScaleWidth      =   8475
      TabIndex        =   3
      Top             =   240
      Width           =   8535
   End
   Begin VB.CommandButton cmdlookupprice 
      Caption         =   "Look up aircraft price"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdsortbyprice 
      Caption         =   "Sort by price"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdreadprint 
      Caption         =   "REad and print pricing information"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblauthor 
      BackColor       =   &H00FF8080&
      Caption         =   "VB Design by Kerry R. O'Neill 10/24/2003"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label lblpricing 
      BackColor       =   &H00FF8080&
      Caption         =   "All Prices are in Millions (US)"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   4920
      Width           =   3735
   End
End
Attribute VB_Name = "AircraftPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this form is to inform the buyer about the pricing
'of each Boeing Aircraft. These prices are the base price and do
'not include any special modifications or upgrades a buyer may wish
'Boeing to include in the aircraft before delivery.


Option Explicit
Dim N(1 To 8) As String 'name of aircraft
Dim Pr(1 To 8) As Single 'price of aircraft (in millions)
Dim i As Integer
Public strpath As String

Private Sub cmdlookupprice_Click() 'searches for aircraft price based on the name user inputs
    pbxresults2.Cls
    Dim Z As String
    Dim done As Boolean
    Z = InputBox("Enter Name", "Aircraft Pricing")
    done = False
    i = 0
    Do Until done Or i = 8  'searches until match is found
       i = i + 1 'moves on to next name
       If Z = N(i) Then done = True
    Loop
    If done Then 'prints what is found
        pbxresults2.Print "Name", Tab(15); "Price"
        pbxresults2.Print "-----------------------------------------"
        pbxresults2.Print 'provides a space between heading and data
        pbxresults2.Print N(i), Tab(15); Pr(i)
    Else
        MsgBox ("Aircraft Not Found") 'alerts user that entry is not valid
    End If
End Sub

Private Sub cmdreadprint_Click() 'reads and prints a text file into visual basic
    pbxresults2.Cls
    Open strpath & "boeingprices.txt" For Input As #1 'opens text file for vb project
    pbxresults2.Print "Name", Tab(15); "Price"
    pbxresults2.Print "--------------------------------------------"
    pbxresults2.Print 'provides a space between heading and data
    For i = 1 To 8
        Input #1, N(i), Pr(i) 'inputs file into two parallel arrays
        pbxresults2.Print N(i), Tab(15); Pr(i) 'prints results
    Next i
    Close #1
End Sub

Private Sub cmdmainmenu_Click() 'returns user to main menu
    AircraftPrice.Hide
    MainMenu.Show
    
End Sub

Private Sub cmdsortbyprice_Click()
    Dim pass As Integer
    Dim X As Integer
    Dim temp1 As String 'temp files for use in sorting sequences
    Dim temp2 As Single
    X = 8
    For pass = 1 To (X - 1) 'bubble sort to make sure data is in correct specified order
        For i = 1 To (X - pass)
            If Pr(i) < Pr(i + 1) Then
                temp1 = N(i)
                N(i) = N(i + 1)
                N(i + 1) = temp1
                temp2 = Pr(i)
                Pr(i) = Pr(i + 1)
                Pr(i + 1) = temp2
            End If
        Next i
    Next pass
    pbxresults2.Cls
    pbxresults2.Print "Name", Tab(15); "Price"
    pbxresults2.Print "--------------------------------------------"
    pbxresults2.Print 'provides a space between heading and data
    For i = 1 To 8
        pbxresults2.Print N(i), Tab(15); Pr(i) 'prints sorted data
    Next i
    
        
End Sub

Private Sub Form_Load() 'creates a strpath so the file can be opened after being moved to different folders
    strpath = "N:\CS130\handin\KRONEILL\"
End Sub
