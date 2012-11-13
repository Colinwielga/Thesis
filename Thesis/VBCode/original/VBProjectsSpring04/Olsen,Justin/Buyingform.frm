VERSION 5.00
Begin VB.Form Buyingform 
   BackColor       =   &H00404040&
   Caption         =   "Form2"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form2"
   ScaleHeight     =   8580
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdgolink 
      BackColor       =   &H0000FF00&
      Caption         =   "Go to my links!"
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdpic6 
      BackColor       =   &H00FF0000&
      Caption         =   "Picture"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmd12 
      BackColor       =   &H000000C0&
      Caption         =   "12' Round-About"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdpic7 
      BackColor       =   &H00FF0000&
      Caption         =   "Picture"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton cmdpic5 
      BackColor       =   &H00FF0000&
      Caption         =   "Picture"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdpic4 
      BackColor       =   &H00FF0000&
      Caption         =   "Picture"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdpic3 
      BackColor       =   &H00FF0000&
      Caption         =   "Picture"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdpic2 
      BackColor       =   &H00FF0000&
      Caption         =   "Picture"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdpic1 
      BackColor       =   &H00FF0000&
      Caption         =   "Picture"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmddone 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdgoback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go back to the beginning."
      Height          =   735
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdall 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to see a list of all canoes and their respective prices, tax included."
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      Height          =   3975
      Left            =   1560
      ScaleHeight     =   3915
      ScaleWidth      =   9075
      TabIndex        =   7
      Top             =   960
      Width           =   9135
   End
   Begin VB.CommandButton cmd19 
      BackColor       =   &H000000C0&
      Caption         =   "19' Northwoods"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmd18 
      BackColor       =   &H000000C0&
      Caption         =   "18' Hauler"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmd17 
      BackColor       =   &H000000C0&
      Caption         =   "17' Streamline"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmd16 
      BackColor       =   &H000000C0&
      Caption         =   "16' Skipper"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmd15 
      BackColor       =   &H000000C0&
      Caption         =   "15' Mini"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmd14 
      BackColor       =   &H000000C0&
      Caption         =   "14' Cruiser"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   $"Buyingform.frx":0000
      Height          =   855
      Left            =   1920
      TabIndex        =   22
      Top             =   6120
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Click on the blue tabs to see that canoes picture!"
      Height          =   495
      Left            =   3360
      TabIndex        =   20
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Designed by: Justin Olsen"
      Height          =   255
      Left            =   9000
      TabIndex        =   11
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the red tabs below to see the canoes cooresponding price, and a little information about it."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Buyingform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Path As String
'Purpose = This form is here to inform the user about different canoes they might want to buy.
'Project Name= Visual Basic Canoe Project("M:\CS130\CanoeProject")
'Form Name= Form 1 ("M:\CS130\CanoeProject\Project1.vbp\Buyingform.frm")

Private Sub cmd12_Click()
picResults.Cls
picResults.Print "The Roundabout is a small, light, and simple, double paddle canoe for one person."
picResults.Print "The price for this beauty is $1200."
End Sub

Private Sub cmd14_Click()
picResults.Cls
picResults.Print "The Cruiser is a wonderful single-paddler craft if you are a little heavier or want to take your dog along with you occasionally."
picResults.Print "The price for this beauty is $1550."
End Sub

Private Sub cmd15_Click()
picResults.Cls
picResults.Print "The Mini is loved by fisherman as well as solo trippers."
picResults.Print "The price for the beauty is $1700."
End Sub

Private Sub cmd16_Click()
picResults.Cls
picResults.Print "The Skipper is a great little two person canoe.  Not quite suitable for long trips, but great for crusing around for a day."
picResults.Print "The price for this beauty is $2100."
End Sub

Private Sub cmd17_Click()
picResults.Cls
picResults.Print "The Streamline is the fastest boat we produce, no one paddling will be faster!"
picResults.Print "The price for this beauty is $2500."
End Sub

Private Sub cmd18_Click()
picResults.Cls
picResults.Print "The Hauler is our premier tripping vessel, there is enough room for two paddlers and a weeks worth of supplies."
picResults.Print "The price for this beauty is $3000."
End Sub

Private Sub cmd19_Click()
picResults.Cls
picResults.Print "The Northwoods is stricltly for HEAVY LOADS.  You can fit upto 4 adults and 3 kids in to this one!"
picResults.Print "The price for this beauty is $3500."
End Sub

Private Sub cmdall_Click()
Dim Canoes(1 To 7) As String, Prices(1 To 7) As String, J As Integer, Tax As Single, Total As Double
    'Prepare the file to be read
    Open Path & "canoearray.txt" For Input As #1
    picResults.Cls
    picResults.Print "Canoes"; Tab(20); "Prices"; Tab(40); "Tax"; Tab(60); "Total"
    picResults.Print "**************************************************************************************"
    For J = 1 To 7
        Input #1, Canoes(J), Prices(J)
        Tax = Prices(J) * 0.07
        Total = Prices(J) + Tax
        picResults.Print Canoes(J); Tab(20); FormatCurrency(Prices(J)); Tab(40); FormatCurrency(Tax); Tab(60); FormatCurrency(Total)
    Next J
    Close #1
End Sub

Private Sub cmddone_Click()
End
End Sub

Private Sub cmdgoback_Click()
Form1.Show
Buyingform.Hide
End Sub

Private Sub cmdgolink_Click()
Buyingform.Hide
Form2.Show
End Sub

Private Sub cmdpic1_Click()
Picture1.Show
Buyingform.Hide
End Sub

Private Sub cmdpic2_Click()
Picture2.Show
Buyingform.Hide
End Sub

Private Sub cmdpic3_Click()
Picture3.Show
Buyingform.Hide
End Sub

Private Sub cmdpic4_Click()
Picture5.Show
Buyingform.Hide
End Sub

Private Sub cmdpic5_Click()
Picture4.Show
Buyingform.Hide
End Sub

Private Sub cmdpic6_Click()
Picture6.Show
Buyingform.Hide
End Sub

Private Sub cmdpic7_Click()
Picture7.Show
Buyingform.Hide
End Sub

Private Sub Form_Load()
Path = "M:\CS130\CanoeProject\"
End Sub

