VERSION 5.00
Begin VB.Form frmMaps 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Republic of Srpska, election 2006"
   ClientHeight    =   9840
   ClientLeft      =   1305
   ClientTop       =   855
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   11025
   Begin VB.CommandButton cmdCounty2 
      BackColor       =   &H008080FF&
      Caption         =   "County 2 map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCounty3 
      BackColor       =   &H008080FF&
      Caption         =   "County 3 map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCounty4 
      BackColor       =   &H008080FF&
      Caption         =   "County 4 map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCounty5 
      BackColor       =   &H008080FF&
      Caption         =   "County 5 map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCounty6 
      BackColor       =   &H008080FF&
      Caption         =   "County 6 map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Back to main page"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCounty1 
      BackColor       =   &H008080FF&
      Caption         =   "County 1 map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   5760
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   4
      Top             =   5760
      Width           =   5055
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   5760
      ScaleHeight     =   3675
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   2040
      Width           =   5055
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   120
      ScaleHeight     =   7515
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   2040
      Width           =   5655
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         ScaleHeight     =   375
         ScaleWidth      =   15
         TabIndex        =   2
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.CommandButton cmdMapBH 
      BackColor       =   &H0080C0FF&
      Caption         =   "View map of Bosnia and Herzegovina and Republic of Srpska"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "frmMaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This part of program presents how RS is organized in 6 different counties
'and people from which town voting in which county. In this part is presented
'map of RS and map of every county. This help guest that have visual understanding for
'model of voting in Republic of Srpska

'This part define Global Values for form Maps
Dim CityCode(1 To 15) As String, CityName(1 To 15) As String

Private Sub cmdCounty1_Click()

'This part define values for private sub
Dim Ctr As Integer

'Whit this option program clean output
picResults1.Cls
picResults2.Cls
picResults3.Cls

'option organize output
picResults3.Print "COUNTY 1"
picResults3.Print
picResults3.Print "City code"; Tab(18); "City name"

'option open data for this private sub
Open App.Path & "\Cities_County_1.txt" For Input As #1

'start array and read information from data define as input #1
Ctr = 0
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, CityCode(Ctr), CityName(Ctr)
picResults3.Print CityCode(Ctr); Tab(18); CityName(Ctr) ' print information in output #3
Loop

'Load picture and print it in output #1
picResults1.Picture = LoadPicture(App.Path & "\" & "Copy (2) of Republic of Srpska counties.jpg")

'Load picture and print it in output #2
picResults2.Picture = LoadPicture(App.Path & "\" & "County 1.bmp")

Close #1
End Sub

Private Sub cmdCounty2_Click()
'This part define values for private sub
Dim Ctr As Integer

'Whit this option program clean output
picResults1.Cls
picResults2.Cls
picResults3.Cls

'option organize output
picResults3.Print "COUNTY 2"
picResults3.Print
picResults3.Print "City code"; Tab(18); "City name"

Open App.Path & "\Cities_County_2.txt" For Input As #1


'start array and read information from data define as input #1
Ctr = 0
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, CityCode(Ctr), CityName(Ctr)
picResults3.Print CityCode(Ctr); Tab(18); CityName(Ctr) ' print information in output #3
Loop

'Load picture and print it in output #1
picResults1.Picture = LoadPicture(App.Path & "\" & "Copy (2) of Republic of Srpska counties.jpg")

'Load picture and print it in output #2
picResults2.Picture = LoadPicture(App.Path & "\" & "County 2.bmp")

Close #1

End Sub

Private Sub cmdCounty3_Click()

'This part define values for private sub
Dim Ctr As Integer

'Whit this option program clean output
picResults1.Cls
picResults2.Cls
picResults3.Cls

'option organize output
picResults3.Print "COUNTY 3"
picResults3.Print
picResults3.Print "City code"; Tab(18); "City name"

Open App.Path & "\Cities_County_3.txt" For Input As #1


'start array and read information from data define as input #1
Ctr = 0
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, CityCode(Ctr), CityName(Ctr)
picResults3.Print CityCode(Ctr); Tab(18); CityName(Ctr) ' print information in output #3
Loop

'Load picture and print it in output #1
picResults1.Picture = LoadPicture(App.Path & "\" & "Copy (2) of Republic of Srpska counties.jpg")

'Load picture and print it in output #2
picResults2.Picture = LoadPicture(App.Path & "\" & "County 3.bmp")

Close #1

End Sub

Private Sub cmdCounty4_Click()

'This part define values for private sub
Dim Ctr As Integer

'Whit this option program clean output
picResults1.Cls
picResults2.Cls
picResults3.Cls

'option organize output
picResults3.Print "COUNTY 4"
picResults3.Print
picResults3.Print "City code"; Tab(18); "City name"

Open App.Path & "\Cities_County_4.txt" For Input As #1


'start array and read information from data define as input #1
Ctr = 0
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, CityCode(Ctr), CityName(Ctr)
picResults3.Print CityCode(Ctr); Tab(18); CityName(Ctr) ' print information in output #3
Loop

'Load picture and print it in output #1
picResults1.Picture = LoadPicture(App.Path & "\" & "Copy (2) of Republic of Srpska counties.jpg")

'Load picture and print it in output #2
picResults2.Picture = LoadPicture(App.Path & "\" & "County 4.bmp")

Close #1

End Sub

Private Sub cmdCounty5_Click()
'This part define values for private sub
Dim Ctr As Integer

'Whit this option program clean output
picResults1.Cls
picResults2.Cls
picResults3.Cls

'option organize output
picResults3.Print "COUNTY 5"
picResults3.Print
picResults3.Print "City code"; Tab(18); "City name"

Open App.Path & "\Cities_County_5.txt" For Input As #1


'start array and read information from data define as input #1
Ctr = 0
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, CityCode(Ctr), CityName(Ctr)
picResults3.Print CityCode(Ctr); Tab(18); CityName(Ctr) ' print information in output #3
Loop

'Load picture and print it in output #1
picResults1.Picture = LoadPicture(App.Path & "\" & "Copy (2) of Republic of Srpska counties.jpg")

'Load picture and print it in output #2
picResults2.Picture = LoadPicture(App.Path & "\" & "County 5.bmp")

Close #1

End Sub

Private Sub cmdCounty6_Click()
'This part define values for private sub
Dim Ctr As Integer

'Whit this option program clean output
picResults1.Cls
picResults2.Cls
picResults3.Cls

'option organize output
picResults3.Print "COUNTY 6"
picResults3.Print
picResults3.Print "City code"; Tab(18); "City name"

Open App.Path & "\Cities_County_6.txt" For Input As #1


'start array and read information from data define as input #1
Ctr = 0
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, CityCode(Ctr), CityName(Ctr)
picResults3.Print CityCode(Ctr); Tab(18); CityName(Ctr) ' print information in output #3
Loop

'Load picture and print it in output #1
picResults1.Picture = LoadPicture(App.Path & "\" & "Copy (2) of Republic of Srpska counties.jpg")

'Load picture and print it in output #2
picResults2.Picture = LoadPicture(App.Path & "\" & "County 6.bmp")

Close #1

End Sub

Private Sub cmdMapBH_Click()

'Whit this option program clean output
picResults1.Cls
picResults2.Cls
picResults3.Cls

'Load picture and print it in output #1
picResults1.Picture = LoadPicture(App.Path & "\" & "Copy of B&H map.jpg")

'Load picture and print it in output #2
picResults2.Picture = LoadPicture(App.Path & "\" & "Copy of Republic of Srpska counties.jpg")

' print information in output #3
picResults3.Print
picResults3.Print "Republic of Srpska is seperated in 6 counties:"
picResults3.Print
picResults3.Print "County 1, region Krajina, 10 cities"
picResults3.Print "County 2, region Banja Luka, 11 cities"
picResults3.Print "County 3, region Posavina, 7 cities"
picResults3.Print "County 4, region Semberija & Majevica, 7 cities"
picResults3.Print "County 5, region Podrinje & Romanija , 15 cities"
picResults3.Print "County 6, region Herzegovina, 13 cities"

'Close #1
'Close #2

End Sub

Private Sub Command2_Click()
frmMainPage.Show
frmMaps.Hide
End Sub
