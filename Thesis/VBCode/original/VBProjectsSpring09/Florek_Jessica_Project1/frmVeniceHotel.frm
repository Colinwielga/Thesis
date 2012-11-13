VERSION 5.00
Begin VB.Form frmVeniceHotel 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBook 
      BackColor       =   &H00C0C000&
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFC0&
      FillColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   4755
      TabIndex        =   5
      Top             =   4680
      Width           =   4815
   End
   Begin VB.OptionButton optHostel 
      Caption         =   "Option3"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optMedium 
      Caption         =   "Option2"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optExpensive 
      Caption         =   "Option1"
      Height          =   195
      Left            =   5280
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C000&
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblHostel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Al Due Leoncini"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   5760
      TabIndex        =   10
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblMedHotel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Albergo Doni"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblExpensiveHotel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Al Ponte Antico"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Book A Hotel In Venice"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmVeniceHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmVeniceHotel
'Jessica Florek
'Written: 3/12/09
'Objective: This form has a variety of options for hotels in Venice that the user can choose
'and will be added to their budgets.



Option Explicit

Private Sub cmdBack_Click()
picResults2.Cls
frm.VeniceHotel.Hide
frmVenice.Show

End Sub

Private Sub cmdBook_Click()
Dim numbernights As Integer, found As Boolean
'calculates cost of staying that they hotel of users choice
numbernights = InputBox("How many nights will you be staying?")
duration = duration - numbernights
If duration < 0 Then
    MsgBox ("You have exceeded your original trip duration by " & Abs(duration) & " days.")
End If

If optExpensive = True Then
        Venicehotelcost = (250 * numbernights)
        budget = budget - Venicehotelcost
        found = True
ElseIf optMedium = True Then
        Venicehotelcost = (60 * numbernights)
        budget = budget - Venicehotelcost
        found = True
ElseIf optHostel = True Then
        Venicehotelcost = (15 * numbernights)
        budget = budget - Venicehotelcost
        found = True
ElseIf (Not found) Then
        MsgBox ("Please select a hotel by clicking on the corresponding bubble.")
End If

venice = True
'makes the information relating to venice displayed during budget summary
frmVeniceHotel.Hide
frmVenice.Show
End Sub

'each option button that is selected displays a hotel picture and corresponding information into picture boxes
Private Sub optExpensive_Click()
picResults.Picture = LoadPicture(App.Path & "\VenicePics\VeniceHotelExpensive.jpg")
picResults2.Print "Al Ponte Antico *****"
picResults2.Print "Luxurious and Elegant."
picResults2.Print "Price Per Night is $250."
picResults2.Print

End Sub

Private Sub optHostel_Click()
picResults.Picture = LoadPicture(App.Path & "\VenicePics\VeniceHostel.jpg")
picResults2.Print "Al Due Leoncini"
picResults2.Print "Classy yet economical!"
picResults2.Print "Price Per Night is $15."
picResults2.Print
End Sub

Private Sub optMedium_Click()
picResults.Picture = LoadPicture(App.Path & "\VenicePics\VeniceHotelMedium.jpg")
picResults2.Print "Albergo Doni ***."
picResults2.Print "Traveler's Favorite!"
picResults2.Print "Price Per Night is $60."
picResults2.Print
End Sub
