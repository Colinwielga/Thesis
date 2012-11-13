VERSION 5.00
Begin VB.Form frmLondonHotel 
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBook 
      BackColor       =   &H80000013&
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2355
      ScaleWidth      =   5475
      TabIndex        =   9
      Top             =   5040
      Width           =   5535
   End
   Begin VB.OptionButton optHostel 
      Caption         =   "Option3"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optMedium 
      Caption         =   "Option2"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optExpensive 
      Caption         =   "Option1"
      Height          =   195
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000003&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000013&
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label lblHostel 
      BackColor       =   &H80000003&
      Caption         =   "Ashlee House Hostel"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   6000
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblMedHotel 
      BackColor       =   &H80000003&
      Caption         =   "Europa Gatwick Hotel"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblExpensiveHotel 
      BackColor       =   &H80000003&
      Caption         =   "Draycott Hotel       "
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000003&
      Caption         =   "Book A Hotel In London!"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmLondonHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmLondonHotel
'Jessica Florek
'Written: 3/6/09
'Objective: This form has a variety of options for hotels in London that the user can choose
'and will be added to their budgets.


Option Explicit

Private Sub cmdBack_Click()
picResults2.Cls

frmLondonHotel.Hide
frmLondon.Show

End Sub

Private Sub cmdBook_Click()
Dim numbernights As Integer, found As Boolean
'books hotel by adding the cost of the hotel times the number of nights to the londonhotelcost which will be used during the budget summary, and it subtracts this expense from the budget
numbernights = InputBox("How many nights will you be staying?")
duration = duration - numbernights
'if the user books more nights than they had intented for their duration this message pops up informing them of this
If duration < 0 Then
    MsgBox ("You have exceeded your original trip duration by " & Abs(duration) & " days.")
End If

'depending on the hotel that was selected via an option button, the costs are calculated
If optExpensive = True Then
        Londonhotelcost = (305 * numbernights)
        budget = budget - Londonhotelcost
        found = True
ElseIf optMedium = True Then
        Londonhotelcost = (40 * numbernights)
        budget = budget - Londonhotelcost
        found = True
ElseIf optHostel = True Then
        Londonhotelcost = (15 * numbernights)
        budget = budget - Londonhotelcost
        found = True
ElseIf (Not found) Then
        MsgBox ("Please select a hotel by clicking on the corresponding bubble.")
End If

london = True
'this causes the london summaries to be shown in the overall budget summaries
frmLondonHotel.Hide
frmLondon.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub


'for each option button that is clicked, the hotel information and a picture of the hotel room is displayed in picture boxes

Private Sub optExpensive_Click()
picResults.Picture = LoadPicture(App.Path & "\LondonPics\LondonHotelExpensiveRoom2.jpg")
picResults2.Print "The London Draycott Hotel *****"
picResults2.Print "Located in the heart of London, this hotel is a gorgeous retreat."
picResults2.Print "Price Per Night is $305."
picResults2.Print

End Sub

Private Sub optHostel_Click()
picResults.Picture = LoadPicture(App.Path & "\LondonPics\LondonHostel.jpg")
picResults2.Print "Ashlee House Hostel"
picResults2.Print "Cute, modest accomdations!"
picResults2.Print "Price Per Night is $15."
picResults2.Print
End Sub

Private Sub optMedium_Click()
picResults.Picture = LoadPicture(App.Path & "\LondonPics\LondonHotelMediumRoom.jpg")
picResults2.Print "The Europa Gatwick Hotel ***"
picResults2.Print "This reasonably priced hotel is well located for site seeing!"
picResults2.Print "Price Per Night is $40."
picResults2.Print
End Sub
