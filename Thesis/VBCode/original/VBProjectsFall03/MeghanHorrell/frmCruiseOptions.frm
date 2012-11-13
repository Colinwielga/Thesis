VERSION 5.00
Begin VB.Form frmCruiseOptions 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Cruise Options"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSuiteFour 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 4 day cruises sorted by Suite!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   13320
      Width           =   2175
   End
   Begin VB.CommandButton cmdDestinationFour 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 4 day cruises sorted by destination!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   11880
      Width           =   2415
   End
   Begin VB.CommandButton cmdSuiteFive 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 5 day cruises sorted by Suite!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton cmdDestinationFive 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 5 day cruises sorted by destination!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdSuiteSeven 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 7 day cruises sorted by Suite!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdDestinationSeven 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 7 day cruises sorted by destination!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   2415
   End
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   12120
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   12120
      Width           =   2415
   End
   Begin VB.PictureBox pbxResults 
      Height          =   8775
      Left            =   7200
      ScaleHeight     =   8715
      ScaleWidth      =   8475
      TabIndex        =   7
      Top             =   3120
      Width           =   8535
   End
   Begin VB.CommandButton cmdPriceFour 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 4 day cruises sorted by price!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   11880
      Width           =   2535
   End
   Begin VB.CommandButton cmdPriceFive 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 5 day cruises sorted by price"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton cmdPriceSeven 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 7 day cruises sorted by price!"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
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
      Left            =   960
      TabIndex        =   5
      Top             =   10680
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
      Top             =   6600
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
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label lblCruiseOptions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cruise Options"
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
Attribute VB_Name = "frmCruiseOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdDestinationFive_Click()
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 30
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\5 day Cruises.txt" For Input As #1
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If CruiseDestination(i) > CruiseDestination(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdDestinationFour_Click()
Dim CruiseDestination(1 To 18) As String
Dim Suite(1 To 18) As String
Dim Price(1 To 18) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 18
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\4 day Cruises.txt" For Input As #1
    For i = 1 To 18
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If CruiseDestination(i) > CruiseDestination(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 18
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdDestinationSeven_Click()
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 30
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\7 day Cruises.txt" For Input As #1
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If CruiseDestination(i) > CruiseDestination(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub
Private Sub cmdPriceFive_Click()
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 30
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\5 day Cruises.txt" For Input As #1
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdPriceFour_Click()
Dim CruiseDestination(1 To 18) As String
Dim Suite(1 To 18) As String
Dim Price(1 To 18) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 18
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\4 day Cruises.txt" For Input As #1
    For i = 1 To 18
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 18
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturnToHomePage_Click()
    frmDestinations.Hide
    frmHome.Show
End Sub
Private Sub cmdPriceSeven_Click()
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 30
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\7 day Cruises.txt" For Input As #1
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdSuiteFive_Click()
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 30
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\5 day Cruises.txt" For Input As #1
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Suite(i) > Suite(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdSuiteFour_Click()
Dim CruiseDestination(1 To 18) As String
Dim Suite(1 To 18) As String
Dim Price(1 To 18) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 18
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\4 day Cruises.txt" For Input As #1
    For i = 1 To 18
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Suite(i) > Suite(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 18
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdSuiteSeven_Click()
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
N = 30
pbxResults.Cls
pbxResults.Print "Cruise Destination"; Tab(50); "Suite"; Tab(80); "Price"
Open "M:\CS130\Paradise Cruises\7 day Cruises.txt" For Input As #1
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    Close #1
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Suite(i) > Suite(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        pbxResults.Print CruiseDestination(i); Tab(50); Suite(i); Tab(80); FormatCurrency(Price(i))
    Next i
End Sub
