VERSION 5.00
Begin VB.Form frmSpecialOffers 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Special Offers"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturnToHomePage 
      BackColor       =   &H00FF00FF&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   12480
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   12480
      Width           =   2415
   End
   Begin VB.PictureBox pbxResultsFour 
      Height          =   2775
      Left            =   7080
      ScaleHeight     =   2715
      ScaleWidth      =   8355
      TabIndex        =   9
      Top             =   9240
      Width           =   8415
   End
   Begin VB.PictureBox pbxResultsFive 
      Height          =   2655
      Left            =   7080
      ScaleHeight     =   2595
      ScaleWidth      =   8355
      TabIndex        =   8
      Top             =   5640
      Width           =   8415
   End
   Begin VB.PictureBox pbxResultsSeven 
      Height          =   2535
      Left            =   7080
      ScaleHeight     =   2475
      ScaleWidth      =   8355
      TabIndex        =   7
      Top             =   2280
      Width           =   8415
   End
   Begin VB.CommandButton cmdFourDay 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to find out about our special offers for 4 day cruises beginning with the lowest price!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10800
      Width           =   5775
   End
   Begin VB.CommandButton cmdFiveDay 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to find out about our special offers for 5 day cruises beginning with the lowest price!"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   5655
   End
   Begin VB.CommandButton cmdSevenDay 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to find out about our special offers for 7 day cruises beginning with the lowest price!"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label lblFourDay 
      BackColor       =   &H00FFC0C0&
      Caption         =   "4 day cruises"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   5
      Top             =   9360
      Width           =   5535
   End
   Begin VB.Label lblFiveDay 
      BackColor       =   &H00FFC0C0&
      Caption         =   "5 day cruises"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   3
      Top             =   5760
      Width           =   5535
   End
   Begin VB.Label lblSevenDays 
      BackColor       =   &H00FFC0C0&
      Caption         =   "7 day cruises"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label lblSpecialOffers 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Special Offers"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   12015
   End
End
Attribute VB_Name = "frmSpecialOffers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdFiveDay_Click()
    Dim CruiseDestination(1 To 4) As String
    Dim CruiseDate(1 To 4) As String
    Dim Price(1 To 4) As String
    Dim i As Integer
    Dim pass As Integer
    Dim tempPrice As Integer
    Dim tempCruiseDestination As String
    Dim tempCruiseDate As String
    Dim N As Integer
    N = 4
    pbxResultsFive.Print "Cruise Destination and Suite"; Tab(50); "Cruise Date"; Tab(80); "Price"
    Open "M:\CS130\Paradise Cruises\5 day Cruises.txt" For Input As #1
    For i = 1 To 4
        Input #1, CruiseDestination(i), CruiseDate(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempCruiseDate = CruiseDate(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                CruiseDate(i) = CruiseDate(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                CruiseDate(i + 1) = tempCruiseDate
            End If
        Next i
    Next pass
    For i = 1 To 4
        pbxResultsFive.Print CruiseDestination(i); Tab(50); CruiseDate(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdFourDay_Click()
    Dim CruiseDestination(1 To 3) As String
    Dim CruiseDate(1 To 3) As String
    Dim Price(1 To 3) As String
    Dim i As Integer
    Dim pass As Integer
    Dim tempPrice As Integer
    Dim tempCruiseDestination As String
    Dim tempCruiseDate As String
    Dim N As Integer
    N = 3
    pbxResultsFour.Print "Cruise Destination and Suite"; Tab(50); "Cruise Date"; Tab(80); "Price"
    Open "M:\CS130\Paradise Cruises\4 day Cruises.txt" For Input As #1
    For i = 1 To 3
        Input #1, CruiseDestination(i), CruiseDate(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempCruiseDate = CruiseDate(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                CruiseDate(i) = CruiseDate(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                CruiseDate(i + 1) = tempCruiseDate
            End If
        Next i
    Next pass
    For i = 1 To 3
        pbxResultsFour.Print CruiseDestination(i); Tab(50); CruiseDate(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturnToHomePage_Click()
    frmDestinations.Hide
    frmHome.Show
End Sub

Private Sub cmdSevenDay_Click()
    Dim CruiseDestination(1 To 10) As String
    Dim CruiseDate(1 To 10) As String
    Dim Price(1 To 10) As Integer
    Dim i As Integer
    Dim pass As Integer
    Dim tempPrice As Integer
    Dim tempCruiseDestination As String
    Dim tempCruiseDate As String
    Dim tempSuite As String
    Dim N As Integer
    N = 10
    pbxResultsSeven.Print "Cruise Destination and Suite"; Tab(50); "Cruise Start Date"; Tab(80); "Price"
    Open "M:\CS130\Paradise Cruises\7 day Cruises.txt" For Input As #1
    For i = 1 To 10
        Input #1, CruiseDestination(i), CruiseDate(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempCruiseDate = CruiseDate(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                CruiseDate(i) = CruiseDate(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                CruiseDate(i + 1) = tempCruiseDate
            End If
        Next i
    Next pass
    For i = 1 To 10
        pbxResultsSeven.Print CruiseDestination(i); Tab(50); CruiseDate(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

