VERSION 5.00
Begin VB.Form frmPlace 
   BackColor       =   &H8000000D&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   ScaleHeight     =   8520
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0080FF80&
      Caption         =   "Some Help Deciding"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   12
      Top             =   6000
      Width           =   3135
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Its Now Time To Decide a Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      TabIndex        =   11
      Top             =   7440
      Width           =   3735
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      Height          =   3015
      Left            =   4920
      ScaleHeight     =   2955
      ScaleWidth      =   5715
      TabIndex        =   10
      Top             =   2760
      Width           =   5775
   End
   Begin VB.CommandButton cmdPlaces 
      BackColor       =   &H00C0FFC0&
      Caption         =   "         Click HERE For           Current Weather"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   5565
      Left            =   2040
      Picture         =   "frmPlace.frx":0000
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblPlaces 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Possible Destinations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label lblBah 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Bahamas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label lblPan 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "Panama City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label lblAtl 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Atlanta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lblMiami 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "Miami"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblSan 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "San Diego"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblLas 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "Las Vegas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblCan 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Cancun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblWhere 
      Alignment       =   2  'Center
      Caption         =   "Where To Go?"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   36
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmPlace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Weather page'
'This page exibits color and skill of VB'
'It opens up files, usedes MSGBox's and many other things'
' it reads a text file'
'october 14th 2009'
'blake abuer'

Option Explicit
Dim City(1 To 30) As String, Temp(1 To 100) As Single, Weather(1 To 30) As String
Dim Ctr As Integer



'Quit Button'
Private Sub cmdEnd_Click()
    End
End Sub

'creating variables so i can use a input box'

Private Sub cmdHelp_Click()
Dim Destination As String

Ctr = 0
'opening up a file'
Open App.Path & "\cities1.txt" For Input As #1



Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, City(Ctr)
Loop
Close #1
    'creating an inputbox'
    Destination = InputBox(" What City Were You Considering Going To For Spring Break?(Make Sure to Capitalize the first letter of the city)", "Just A Little Help")
      'making an if then statments with inputbox responses'
      If Destination = "Cancun" Then
            MsgBox "Cancun is a great place for you if you are looking for hot sun, crazy parties and late nights. The drinking age is under 21, so that is a plus for some people. ", , "Good Choice"
     ElseIf Destination = "Las Vegas" Then
            MsgBox "Las Vegas is a great place for a person willing to spend a little more money on accommodations, and wants to gamble. Nice weather, and 1000's of things to do.", , "Good Choice"
     ElseIf Destination = "San Diego" Then
            MsgBox "San Diego is a great place for golfers and surfers. Great family destination with great sea food.", , "Good Choice"
     ElseIf Destination = "Miami" Then
            MsgBox "Miami is the place to be seen if you are rich and famous. If you are neither one of those it's still ok. Miami has great beaches and beautiful people.", , "Good Choice"
     ElseIf Destination = "Atlanta" Then
            MsgBox "Atlanta is a great place to go if you're looking for some great food, and shopping opportunities. It is a very large urban city with many night clubs.", , "Good Choice"
     ElseIf Destination = "Panama City" Then
            MsgBox "Panama City is the Unites States version of Cancun. It is a great place to party, and have fun with friends. Don't forget the drinking age here is 21.", , "Good Choice"
     ElseIf Destination = "Bahamas" Then
            MsgBox "Great beaches, great weather, truly a tropical paradise. Expensive, but worth it for the weather.", , "Good Choice"
     'the else statemnet to catch errors'
     Else
            MsgBox "Not a possible Destination, or you did not capitalize correctly", , "Alert!!!"
      End If
      
cmdNext.Visible = True
cmdHelp.Visible = True
      
        
End Sub
'hid and show frm's'
Private Sub cmdNext_Click()
    frmFlights.Show
    frmPlace.Hide
    
  
End Sub

Private Sub cmdPlaces_Click()


Ctr = 0
'openning an array'
Open App.Path & "\cities1.txt" For Input As #1


picResults.Print "City"; , , ; "Temperature in Degrees"; "   "; "What Its Like Outside"
picResults.Print "***************************************************************************************"
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, City(Ctr), Temp(Ctr), Weather(Ctr)
'printing my array'
picResults.Print City(Ctr); , , ; Temp(Ctr); , , ; Weather(Ctr)

Loop


Close #1


'changing a frm'
cmdPlaces.Visible = False
cmdNext.Visible = False

End Sub


