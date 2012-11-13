VERSION 5.00
Begin VB.Form frmTeam 
   BackColor       =   &H00400000&
   Caption         =   "Hendrick Motorsports"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13815
   LinkTopic       =   "Form4"
   ScaleHeight     =   8685
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display chosen driver"
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.PictureBox picResults1 
      Height          =   2175
      Left            =   600
      ScaleHeight     =   2115
      ScaleWidth      =   4755
      TabIndex        =   4
      Top             =   6120
      Width           =   4815
   End
   Begin VB.PictureBox picResults 
      Height          =   4815
      Left            =   5880
      ScaleHeight     =   4755
      ScaleWidth      =   7755
      TabIndex        =   3
      Top             =   2280
      Width           =   7815
   End
   Begin VB.TextBox txtDriver 
      Height          =   735
      Left            =   6600
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmTeam.frx":0000
      Height          =   1215
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   3645
      Left            =   360
      Picture         =   "frmTeam.frx":008D
      Top             =   2040
      Width           =   5370
   End
End
Attribute VB_Name = "frmTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Introduction to NASCAR
'Form Team
'Colin Roberts and Luke Hommerding
'Written 10/18/09
'Purpose is to display drivers based on a number entered into a text box and using select case
Option Explicit 'declare variables
Dim Driver As Single
'when a value is entered into a text box the corresponding picture and bio are displayed
'on the screen
Private Sub cmdDisplay_Click()
    Driver = txtDriver.Text
    Select Case Driver
        Case Is = 1
            picResults.Cls 'clears picture boxes before new data is entered
            picresults1.Cls
            'prints bio of driver and picture of driver
            picResults.Picture = LoadPicture(App.Path & "\JeffGordon24Car1.jpg")
            picresults1.Print " Jeffrey Michael Gordon"
            picresults1.Print
            picresults1.Print "- Born August 4th 1971"
            picresults1.Print "- Hometown: Vallejo, California"
            picresults1.Print "- Started racing age 5"
            picresults1.Print "- Has won 4 Sprint Cup Championships"
            picresults1.Print "- Currently in Third place in the 2009 Sprint Cup."
        Case Is = 2
            picResults.Cls
            picresults1.Cls
            'prints bio of driver and picture of driver
            picResults.Picture = LoadPicture(App.Path & "\JimmieJohnson48Car.jpg")
            picresults1.Print " Jimmie Kenneth Johnson"
            picresults1.Print
            picresults1.Print "- Born September 17th 1975"
            picresults1.Print "- Hometown: El Cajon, California"
            picresults1.Print "- Won his first championship on a dirt bike"
            picresults1.Print "- Has won 3 consecutive Sprint Cup Championships"
            picresults1.Print "- Currently in First place in the 2009 Sprint Cup."
        Case Is = 3
            picResults.Cls
            picresults1.Cls
            'prints picture and bio of driver
            picResults.Picture = LoadPicture(App.Path & "\MarkMartin5Car.jpg")
            picresults1.Print " Mark Anthony Martin"
            picresults1.Print
            picresults1.Print "- Born January 5th 1959"
            picresults1.Print "- Hometown: Batesville, Arkansas"
            picresults1.Print "- Currently oldest driver in NASCAR"
            picresults1.Print "- Oldest driver to win a Sprint Cup race"
            picresults1.Print "- Currently in Second place in the 2009 Sprint Cup."
        Case Is = 4
            picResults.Cls
            picresults1.Cls
            'prints picture and bio of the driver
            picResults.Picture = LoadPicture(App.Path & "\DaleEarnhardtJr88Car.jpg")
            picresults1.Print Right("Ralph Dale Earnhardt Jr.", 19)
            picresults1.Print
            picresults1.Print "- Born October 10th 1974"
            picresults1.Print "- Hometown: Kannapolis, North Carolina"
            picresults1.Print "- Son of Legendary NASCAR driver Dale Earnhardt Sr."
            picresults1.Print "- Has yet to win a Sprint Cup Championship"
            picresults1.Print "- Currently in 22nd in the 2009 Sprint Cup."
        'displays false info if the number in the text box is not valid
        Case Else
            picResults.Cls
            picresults1.Cls
            picResults.Picture = LoadPicture(App.Path & "\redflag.jpg")
            picresults1.Print "Sorry, the race been stopped due to an invalid entry."
    End Select
        
End Sub
'returns to main menu
Private Sub cmdReturn_Click()
    frmMain.Show
    frmTeam.Hide
End Sub

