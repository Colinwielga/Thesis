VERSION 5.00
Begin VB.Form frmItinerary 
   BackColor       =   &H0000C000&
   Caption         =   "Display of Itinerary Options"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbxResultsAllOptions 
      Height          =   6015
      Left            =   12240
      ScaleHeight     =   5955
      ScaleWidth      =   5595
      TabIndex        =   12
      Top             =   3840
      Width           =   5655
   End
   Begin VB.TextBox txtDays 
      Height          =   975
      Left            =   3600
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturnToHome 
      BackColor       =   &H00FF00FF&
      Caption         =   "Return to Home Page"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10920
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10920
      Width           =   2175
   End
   Begin VB.TextBox txtDestination 
      Height          =   975
      Left            =   3600
      TabIndex        =   3
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox txtPrice 
      Height          =   975
      Left            =   3600
      TabIndex        =   2
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FF00FF&
      Caption         =   "Display your Itinterary Options"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   2175
   End
   Begin VB.PictureBox pbxResultsItinerary 
      Height          =   6015
      Left            =   6240
      ScaleHeight     =   5955
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   3840
      Width           =   5415
   End
   Begin VB.Label lblAllOptions 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Here are all of the options you have chosen"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12600
      TabIndex        =   14
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label lblCurrentSelection 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Your Current Selection is"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   13
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label lblDays 
      BackColor       =   &H0000C000&
      Caption         =   "Enter the number of days you would like to travel for (either 7, 5 or 4)"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label lblItineraryOptions 
      BackColor       =   &H0000C000&
      Caption         =   "Itinerary Options"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   11295
   End
   Begin VB.Label lblDestination 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Enter a destination that you are looking for"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Enter a Price that you are looking for"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Click this button to display your itinerary options"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   8400
      Width           =   2895
   End
End
Attribute VB_Name = "frmItinerary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdDisplay_Click()
Dim Days As Integer
Dim PriceEntered As Single
Dim DestinationEntered As String
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
Dim i As Integer
Dim Flag As Boolean
Flag = False
Days = txtDays.Text
If Days = 7 Then
    Open "M:\CS130\Paradise Cruises\7 day Cruises.txt" For Input As #1
        For i = 1 To 30
            Input #1, CruiseDestination(i), Suite(i), Price(i)
        Next i
    Close #1
    PriceEntered = txtPrice.Text
    DestinationEntered = txtDestination.Text
    pbxResultsAllOptions.Print "_______________________________________________________________________"
    pbxResultsAllOptions.Print "Cruise Destination You Chose"; Tab(40); "Suite"; Tab(60); "Price of Trip"
        For i = 1 To 30
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsAllOptions.Print CruiseDestination(i); Tab(40); Suite(i); Tab(60); FormatCurrency(Price(i))
            End If
        Next i
        If Flag = False Then
            MsgBox "Sorry, we do not have an itinerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
ElseIf Days = 5 Then
        Open "M:\CS130\Paradise Cruises\5 day Cruises.txt" For Input As #1
            For i = 1 To 30
                Input #1, CruiseDestination(i), Suite(i), Price(i)
            Next i
        Close #1
        PriceEntered = txtPrice.Text
        DestinationEntered = txtDestination.Text
        pbxResultsAllOptions.Print "Cruise Destination You Chose"; Tab(40); "Suite"; Tab(60); "Price of Trip"
        For i = 1 To 30
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsAllOptions.Print CruiseDestination(i); Tab(40); Suite(i); Tab(60); FormatCurrency(Price(i))
            End If
        Next i
        If Flag = False Then
            MsgBox "Sorry, we do not have an itenerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
ElseIf Days = 4 Then
        Open "M:\CS130\Paradise Cruises\4 day Cruises.txt" For Input As #1
            For i = 1 To 18
                Input #1, CruiseDestination(i), Suite(i), Price(i)
            Next i
        Close #1
        PriceEntered = txtPrice.Text
        DestinationEntered = txtDestination.Text
        pbxResultsAllOptions.Print "Cruise Destination You Chose"; Tab(40); "Suite"; Tab(60); "Price of Trip"
        For i = 1 To 18
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsAllOptions.Print CruiseDestination(i); Tab(40); Suite(i); Tab(60); FormatCurrency(Price(i))
            End If
        Next i
        If Flag = False Then
            MsgBox "Sorry, we do not have an itenerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
End If
pbxResultsItinerary.Cls
Days = txtDays.Text
If Days = 7 Then
    Open "M:\CS130\Paradise Cruises\7 day Cruises.txt" For Input As #1
        For i = 1 To 30
            Input #1, CruiseDestination(i), Suite(i), Price(i)
        Next i
    Close #1
    PriceEntered = txtPrice.Text
    DestinationEntered = txtDestination.Text
    pbxResultsAllOptions.Print "_______________________________________________________________________"
    pbxResultsItinerary.Print "Cruise Destination You Chose"; Tab(40); "Suite"; Tab(60); "Price of Trip"
        For i = 1 To 30
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsItinerary.Print CruiseDestination(i); Tab(40); Suite(i); Tab(60); FormatCurrency(Price(i))
            End If
        Next i
        If Flag = False Then
            MsgBox "Sorry, we do not have an itinerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
ElseIf Days = 5 Then
        Open "M:\CS130\Paradise Cruises\5 day Cruises.txt" For Input As #1
            For i = 1 To 30
                Input #1, CruiseDestination(i), Suite(i), Price(i)
            Next i
        Close #1
        PriceEntered = txtPrice.Text
        DestinationEntered = txtDestination.Text
        pbxResultsAllOptions.Print "_______________________________________________________________________"
        pbxResultsItinerary.Print "Cruise Destination You Chose"; Tab(40); "Suite"; Tab(60); "Price of Trip"
        For i = 1 To 30
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsItinerary.Print CruiseDestination(i); Tab(40); Suite(i); Tab(60); FormatCurrency(Price(i))
            End If
        Next i
        If Flag = False Then
            MsgBox "Sorry, we do not have an itenerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
ElseIf Days = 4 Then
        Open "M:\CS130\Paradise Cruises\4 day Cruises.txt" For Input As #1
            For i = 1 To 18
                Input #1, CruiseDestination(i), Suite(i), Price(i)
            Next i
        Close #1
        PriceEntered = txtPrice.Text
        DestinationEntered = txtDestination.Text
        pbxResultsAllOptions.Print "_______________________________________________________________________"
        pbxResultsItinerary.Print "Cruise Destination You Chose"; Tab(40); "Suite"; Tab(60); "Price of Trip"
        For i = 1 To 18
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsItinerary.Print CruiseDestination(i); Tab(40); Suite(i); Tab(60); FormatCurrency(Price(i))
            End If
        Next i
        If Flag = False Then
            MsgBox "Sorry, we do not have an itenerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
End If
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturnToHome_Click()
    frmItinerary.Hide
    frmHome.Show
End Sub

