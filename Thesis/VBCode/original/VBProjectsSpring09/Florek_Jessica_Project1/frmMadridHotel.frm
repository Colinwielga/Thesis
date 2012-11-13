VERSION 5.00
Begin VB.Form frmMadridHotel 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H80000013&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   720
      Width           =   4695
   End
   Begin VB.OptionButton optExpensive 
      Caption         =   "Option1"
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton optMedium 
      Caption         =   "Option2"
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optHostel 
      Caption         =   "Option3"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   4800
      Width           =   5535
   End
   Begin VB.CommandButton cmdBook 
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H80000013&
      Caption         =   "Book A Hotel In Madrid!"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblExpensiveHotel 
      BackColor       =   &H80000013&
      Caption         =   "AC Placio de Retiro Hotel       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblMedHotel 
      BackColor       =   &H80000013&
      Caption         =   "Tryp Ambassador Hotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lblHostel 
      BackColor       =   &H80000013&
      Caption         =   "Cat's Hostel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   5880
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
End
Attribute VB_Name = "frmMadridHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmMadridHotel
'Jessica Florek
'Written: 3/10/09
'Objective: This form has a variety of options for hotels in Madrid that the user can choose
'and will be added to their budgets.



Option Explicit

Private Sub cmdBack_Click()
picResults2.Cls
frmMadridHotel.Hide
frmMadrid.Show
End Sub

Private Sub cmdBook_Click()
Dim numbernights As Integer, found As Boolean
'calculates the total number of nights that have been booked in hotels thus far
numbernights = InputBox("How many nights will you be staying?")
duration = duration - numbernights
'if the user books more nights than they had intented for their duration this message pops up informing them of this
If duration < 0 Then
    MsgBox ("You have exceeded your original trip duration by " & Abs(duration) & " days.")
End If

'depending on the option button selected, the corresponding information is calculated into the budget and added to madridhotelexpense which is later used during budget summary
If optExpensive = True Then
        Madridhotelcost = (100 * numbernights)
        budget = budget - Madridhotelcost
        found = True
ElseIf optMedium = True Then
        Madridhotelcost = (38 * numbernights)
        budget = budget - Madridhotelcost
        found = True
ElseIf optHostel = True Then
        Madridhotelcost = (13 * numbernights)
        budget = budget - Madridhotelcost
        found = True
ElseIf (Not found) Then
        MsgBox ("Please select a hotel by clicking on the corresponding bubble.")
End If

madrid = True

frmMadridHotel.Hide
frmMadrid.Show
End Sub


'when any option button is clicked, the corresponding information and a picture of the hotel appears in picture boxes
Private Sub optExpensive_Click()
picResults.Picture = LoadPicture(App.Path & "\MadridPics\HotelExpensiveRoom.jpg")
picResults2.Print "The AC Placio de Retiro ****"
picResults2.Print "Modern, classy accomodations."
picResults2.Print "Price Per Night is $100."
picResults2.Print
End Sub

Private Sub optHostel_Click()
picResults.Picture = LoadPicture(App.Path & "\MadridPics\HostelRoom.jpg")
picResults2.Print "Cat's Hostel "
picResults2.Print "Warm and welcoming!"
picResults2.Print "Price Per Night is $16."
picResults2.Print
End Sub

Private Sub optMedium_Click()
picResults.Picture = LoadPicture(App.Path & "\MadridPics\HotelMedBed.jpg")
picResults2.Print "Tryp Ambassador ***"
picResults2.Print "Cozy and comfortable!"
picResults2.Print "Price Per Night is $38."
picResults2.Print
End Sub
