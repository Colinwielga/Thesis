VERSION 5.00
Begin VB.Form frmParisHotel 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   720
      Width           =   4695
   End
   Begin VB.OptionButton optExpensive 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton optMedium 
      Caption         =   "Option2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optHostel 
      Caption         =   "Option3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   4800
      Width           =   4815
   End
   Begin VB.CommandButton cmdBook 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Book Hotel"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H0000C000&
      Caption         =   "Book A Hotel In Paris"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblExpensiveHotel 
      BackColor       =   &H0000C000&
      Caption         =   "Hotel Vernet"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblMedHotel 
      BackColor       =   &H0000C000&
      Caption         =   "Hotel Relais du Pre"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblHostel 
      BackColor       =   &H0000C000&
      Caption         =   "Hostel du Paris"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   5880
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
End
Attribute VB_Name = "frmParisHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmParisHotel
'Jessica Florek
'Written: 3/8/09
'Objective: This form has a variety of options for hotels in Paris that the user can choose
'and will be added to their budgets.


Option Explicit

Private Sub cmdBack_Click()
picResults2.Cls
frmParisHotel.Hide
frmParis.Show

End Sub

Private Sub cmdBook_Click()
Dim numbernights As Integer, found As Boolean
'takes user input of number of nights and multiplies that number by the cost per night of the hotel that is selected via option buttons
numbernights = InputBox("How many nights will you be staying?")
duration = duration - numbernights
If duration < 0 Then
    MsgBox ("You have exceeded your original trip duration by " & Abs(duration) & " days.")
End If
'if they exceed the original planned duration of their trip the message pops up to inform them
If optExpensive = True Then
        Parishotelcost = (215 * numbernights)
        budget = budget - Parishotelcost
        found = True
ElseIf optMedium = True Then
        Parishotelcost = (60 * numbernights)
        budget = budget - Parishotelcost
        found = True
ElseIf optHostel = True Then
        Parishotelcost = (20 * numbernights)
        budget = budget - Parishotelcost
        found = True
ElseIf (Not found) Then
        MsgBox ("Please select a hotel by clicking on the corresponding bubble.")
End If

paris = True
'paris data will now be displayed in the budget summary
frmParisHotel.Hide
frmParis.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub


'each option button that is clicked displays corresponding hotel information including a picture of the hotel room

Private Sub optExpensive_Click()
picResults.Picture = LoadPicture(App.Path & "\ParisPics\ParisExpensiveHotelRoom.jpg")
picResults2.Print "Hotel Vernet *****"
picResults2.Print "An elegant hotel located near the Eiffel Tower."
picResults2.Print "Price Per Night is $215."
picResults2.Print


End Sub

Private Sub optHostel_Click()
picResults.Picture = LoadPicture(App.Path & "\ParisPics\ParisHostelRoom.jpg")
picResults2.Print "Hostel du Paris"
picResults2.Print "Modern and comfortable!"
picResults2.Print "Price Per Night is $20."
picResults2.Print
End Sub

Private Sub optMedium_Click()
picResults.Picture = LoadPicture(App.Path & "\ParisPics\ParisMedHotelRoom.jpg")
picResults2.Print "Hotel Relais du Pre ***"
picResults2.Print "Adorable and quaint!"
picResults2.Print "Price Per Night is $60."
picResults2.Print
End Sub

