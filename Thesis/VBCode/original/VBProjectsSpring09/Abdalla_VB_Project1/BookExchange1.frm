VERSION 5.00
Begin VB.Form frmBookswap 
   BackColor       =   &H00000000&
   Caption         =   "Form4"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   FillColor       =   &H000000C0&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form4"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000C0&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortVyCourses 
      BackColor       =   &H0000C000&
      Caption         =   "Sort By Course"
      Height          =   975
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearchByCourse 
      BackColor       =   &H000000C0&
      Caption         =   "Search by Courses"
      Height          =   1095
      Index           =   1
      Left            =   120
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSortByPrice 
      BackColor       =   &H0000C000&
      Caption         =   "Sort by Price"
      Height          =   1095
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSortByCourse 
      Caption         =   "Sort by course"
      Height          =   975
      Left            =   8640
      TabIndex        =   4
      Top             =   11040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearchByCourse 
      Caption         =   "Search by Course"
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   11040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearchTitle 
      BackColor       =   &H000000C0&
      Caption         =   "Search by tittle"
      Height          =   1095
      Left            =   2040
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00008080&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   14955
      TabIndex        =   1
      Top             =   480
      Width           =   15015
   End
   Begin VB.CommandButton cmdSortByBookName 
      BackColor       =   &H0000C000&
      Caption         =   "sort  by book Tittle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Caption         =   "Search book by clicking on the buttons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Sort List by clicking on the Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   8
      Top             =   5040
      Width           =   3855
   End
End
Attribute VB_Name = "frmBookswap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Book exchange'
'Form name: frmBookSwap'
'Author: Bibi Abdalla'
'Date: 3/24/2009'
'Objective: allow buyer to see books registered for sell, sort them or search by a particle order'

Option Explicit

Dim Ctr As Integer
Dim Title(1 To 100) As String
Dim Author(1 To 100) As String
Dim Price(1 To 10) As Double
Dim MarketPrice(1 To 10) As Double
Dim Field(1 To 10) As String
Dim ProfName(1 To 10) As String
Dim CourseName(1 To 10) As String
Dim Location(1 To 10) As String
Dim HolderName(1 To 10) As String
Dim ContactInfo(1 To 10) As String
Dim ISB(1 To 100) As String
Dim Pass As Integer
Dim pos As Integer
'temp ( this is mostly used with Sort buttons'
Dim tempTitle As String
Dim tempAuthor As String
Dim tempPrice As Double
Dim tempMarketPrice As Double
Dim tempField As String
Dim tempProfName As String
Dim tempCourseName As String
Dim tempLocation As String
Dim tempHolderName As String
Dim tempContactInfo As String
Dim tempISB As String
'For UserInputs, for search buttons'
Dim UserInputTitle As String
Dim UserInputCourse As String
Dim UserInputLocation As String
Dim UserInputISB As String
'for Search buttons'
Dim found As Boolean





Private Sub cmdQuit_Click()
frmBookswap.Hide
FrmWelcome.Show

End Sub

Private Sub cmdSearchByCourse_Click(Index As Integer)
'search by course #'
 picResult.Cls
   UserInputCourse = InputBox("Pleaser enter ISB #", "Enter ISB #")
    found = False
    Pass = 0
    Do While ((Not found) And (Pass < Ctr))
    Pass = Pass + 1
    If UserInputCourse = CourseName(Pass) Then
        found = True
  
    End If
    Loop
    If found = True Then
        picResult.Print "Title"; Tab(30); "Author"; Tab(60); "Price"; Tab(75); "MarketPrice"; Tab(95); "Professor Name"; Tab(120); "CourseName"; Tab(140); "Location"; Tab(160); "Holder Name"; Tab(180); "Contact"
        picResult.Print "________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
        picResult.Print Title(Pass); Tab(30); Author(Pass); Tab(60); FormatCurrency(Price(Pass)); Tab(75); FormatCurrency(MarketPrice(Pass)); Tab(95); ISB(Pass); Tab(120); Field(Pass); Tab(140); ProfName(Pass); Tab(160); CourseName(Pass); Tab(190); Location(Pass); Tab(215); HolderName(Pass); Tab(235); ContactInfo(Pass)
    End If
    If found = False Then
        MsgBox "Sorry course not avaliable", , "Not Found"
    End If
End Sub

Private Sub cmdSearchTitle_Click()
'searching by Title'
   picResult.Cls
   UserInputTitle = InputBox("Pleaser enter book title", "Enter Book Title")
    found = False
    Pass = 0
    Do While ((Not found) And (Pass < Ctr))
    Pass = Pass + 1
    If UserInputTitle = Title(Pass) Then
        found = True
  
    End If
    Loop
    If found = True Then
        picResult.Print "Title"; Tab(30); "Author"; Tab(60); "Price"; Tab(75); "MarketPrice"; Tab(95); "Professor Name"; Tab(120); "CourseName"; Tab(140); "Location"; Tab(160); "Holder Name"; Tab(180); "Contact"
        picResult.Print "________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
        picResult.Print Title(Pass); Tab(30); Author(Pass); Tab(60); FormatCurrency(Price(Pass)); Tab(75); FormatCurrency(MarketPrice(Pass)); Tab(95); ISB(Pass); Tab(120); Field(Pass); Tab(140); ProfName(Pass); Tab(160); CourseName(Pass); Tab(190); Location(Pass); Tab(215); HolderName(Pass); Tab(235); ContactInfo(Pass)
    End If
    If found = False Then
        MsgBox "Sorry book not found", , "Not Found"
    End If
  
    
End Sub


Private Sub cmdSortByBookName_Click()
'sort by Tittle'
picResult.Cls
For Pass = 1 To Ctr
    For pos = 1 To Ctr - 1
    If Title(pos) > Title(pos + 1) Then
    tempTitle = Title(pos)
    Title(pos) = Title(pos + 1)
    Title(pos + 1) = tempTitle
    'for author'
    tempAuthor = Author(pos)
    Author(pos) = Author(pos + 1)
    Author(pos + 1) = tempAuthor
    'for price'
    tempPrice = Price(pos)
    Price(pos) = Price(pos + 1)
    Price(pos + 1) = tempPrice
    'for MarketPrice'
    tempMarketPrice = MarketPrice(pos)
    MarketPrice(pos) = MarketPrice(pos + 1)
    MarketPrice(pos + 1) = tempMarketPrice
    'For ISB'
    tempISB = ISB(pos)
    ISB(pos) = ISB(pos + 1)
    ISB(pos + 1) = tempISB
    'for field'
    tempField = Field(pos)
    Field(pos) = Field(pos + 1)
    Field(pos + 1) = tempField
    'for ProfName'
    tempProfName = ProfName(pos)
    ProfName(pos) = ProfName(pos + 1)
    ProfName(pos + 1) = tempProfName
    'for CourseName'
    tempCourseName = CourseName(pos)
    CourseName(pos) = CourseName(pos + 1)
    CourseName(pos + 1) = tempCourseName
    'For location'
    tempLocation = Location(pos)
    Location(pos) = Location(pos + 1)
    Location(pos + 1) = tempLocation
    'for HolderName'
    tempLocation = Location(pos)
    Location(pos) = Location(pos + 1)
    Location(pos + 1) = tempLocation
    'For ContactInfo'
    tempContactInfo = ContactInfo(pos)
    ContactInfo(pos) = ContactInfo(pos + 1)
    ContactInfo(pos + 1) = tempContactInfo
End If
Next pos
Next Pass
'printing'
picResult.Print "Title"; Tab(30); "Author"; Tab(60); "Price"; Tab(75); "MarketPrice"; Tab(95); "Professor Name"; Tab(120); "CourseName"; Tab(140); "Location"; Tab(160); "Holder Name"; Tab(180); "Contact"
For Pass = 1 To Ctr
picResult.Print "________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
picResult.Print Title(Pass); Tab(30); Author(Pass); Tab(60); FormatCurrency(Price(Pass)); Tab(75); FormatCurrency(MarketPrice(Pass)); Tab(95); ProfName(Pass); Tab(120); CourseName(Pass); Tab(140); Location(Pass); Tab(160); HolderName(Pass); Tab(180); ContactInfo(Pass)
Next Pass


End Sub





Private Sub cmdSortByPrice_Click()
'sort by Price'
picResult.Cls
For Pass = 1 To Ctr
    For pos = 1 To Ctr - 1
    If Price(pos) > Price(pos + 1) Then
    tempTitle = Title(pos)
    Title(pos) = Title(pos + 1)
    Title(pos + 1) = tempTitle
    'for author'
    tempAuthor = Author(pos)
    Author(pos) = Author(pos + 1)
    Author(pos + 1) = tempAuthor
    'for price'
    tempPrice = Price(pos)
    Price(pos) = Price(pos + 1)
    Price(pos + 1) = tempPrice
    'for MarketPrice'
    tempMarketPrice = MarketPrice(pos)
    MarketPrice(pos) = MarketPrice(pos + 1)
    MarketPrice(pos + 1) = tempMarketPrice
    'For ISB'
    tempISB = ISB(pos)
    ISB(pos) = ISB(pos + 1)
    ISB(pos + 1) = tempISB
    'for field'
    tempField = Field(pos)
    Field(pos) = Field(pos + 1)
    Field(pos + 1) = tempField
    'for ProfName'
    tempProfName = ProfName(pos)
    ProfName(pos) = ProfName(pos + 1)
    ProfName(pos + 1) = tempProfName
    'for CourseName'
    tempCourseName = CourseName(pos)
    CourseName(pos) = CourseName(pos + 1)
    CourseName(pos + 1) = tempCourseName
    'For location'
    tempLocation = Location(pos)
    Location(pos) = Location(pos + 1)
    Location(pos + 1) = tempLocation
    'for HolderName'
    tempLocation = Location(pos)
    Location(pos) = Location(pos + 1)
    Location(pos + 1) = tempLocation
    'For ContactInfo'
    tempContactInfo = ContactInfo(pos)
    ContactInfo(pos) = ContactInfo(pos + 1)
    ContactInfo(pos + 1) = tempContactInfo
End If
Next pos
Next Pass
'printing'
picResult.Print "Title"; Tab(30); "Author"; Tab(60); "Price"; Tab(75); "MarketPrice"; Tab(95); "Professor Name"; Tab(120); "CourseName"; Tab(140); "Location"; Tab(160); "Holder Name"; Tab(180); "Contact"

For Pass = 1 To Ctr
picResult.Print "________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
picResult.Print Title(Pass); Tab(30); Author(Pass); Tab(60); FormatCurrency(Price(Pass)); Tab(75); FormatCurrency(MarketPrice(Pass)); Tab(95); ProfName(Pass); Tab(120); CourseName(Pass); Tab(140); Location(Pass); Tab(160); HolderName(Pass); Tab(180); ContactInfo(Pass)
Next Pass

End Sub





Private Sub cmdSortVyCourses_Click()
'Sort By Course'
picResult.Cls

For Pass = 1 To Ctr
    For pos = 1 To Ctr - 1
    If CourseName(pos) > CourseName(pos + 1) Then
    tempTitle = Title(pos)
    Title(pos) = Title(pos + 1)
    Title(pos + 1) = tempTitle
    'for author'
    tempAuthor = Author(pos)
    Author(pos) = Author(pos + 1)
    Author(pos + 1) = tempAuthor
    'for price'
    tempPrice = Price(pos)
    Price(pos) = Price(pos + 1)
    Price(pos + 1) = tempPrice
    'for MarketPrice'
    tempMarketPrice = MarketPrice(pos)
    MarketPrice(pos) = MarketPrice(pos + 1)
    MarketPrice(pos + 1) = tempMarketPrice
    'For ISB'
    tempISB = ISB(pos)
    ISB(pos) = ISB(pos + 1)
    ISB(pos + 1) = tempISB
    'for field'
    tempField = Field(pos)
    Field(pos) = Field(pos + 1)
    Field(pos + 1) = tempField
    'for ProfName'
    tempProfName = ProfName(pos)
    ProfName(pos) = ProfName(pos + 1)
    ProfName(pos + 1) = tempProfName
    'for CourseName'
    tempCourseName = CourseName(pos)
    CourseName(pos) = CourseName(pos + 1)
    CourseName(pos + 1) = tempCourseName
    'For location'
    tempLocation = Location(pos)
    Location(pos) = Location(pos + 1)
    Location(pos + 1) = tempLocation
    'for HolderName'
    tempLocation = Location(pos)
    Location(pos) = Location(pos + 1)
    Location(pos + 1) = tempLocation
    'For ContactInfo'
    tempContactInfo = ContactInfo(pos)
    ContactInfo(pos) = ContactInfo(pos + 1)
    ContactInfo(pos + 1) = tempContactInfo
End If
Next pos
Next Pass
'printing'
picResult.Print "Title"; Tab(30); "Author"; Tab(60); "Price"; Tab(75); "MarketPrice"; Tab(95); "Professor Name"; Tab(120); "CourseName"; Tab(140); "Location"; Tab(160); "Holder Name"; Tab(180); "Contact"

For Pass = 1 To Ctr
picResult.Print "________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
picResult.Print Title(Pass); Tab(30); Author(Pass); Tab(60); FormatCurrency(Price(Pass)); Tab(75); FormatCurrency(MarketPrice(Pass)); Tab(95); ProfName(Pass); Tab(120); CourseName(Pass); Tab(140); Location(Pass); Tab(160); HolderName(Pass); Tab(180); ContactInfo(Pass)
Next Pass


End Sub

Private Sub form_load()
Open App.Path & "\booklist3.txt" For Input As #1
Ctr = 0
picResult.Cls
Do Until EOF(1)
Ctr = Ctr + 1
Input #1, Title(Ctr), Author(Ctr), Price(Ctr), MarketPrice(Ctr), ISB(Ctr), Field(Ctr), ProfName(Ctr), CourseName(Ctr), Location(Ctr), HolderName(Ctr), ContactInfo(Ctr)
picResult.Print Title(Ctr), Author(Ctr), Price(Ctr), MarketPrice(Ctr), ISB(Ctr), Field(Ctr), ProfName(Ctr), CourseName(Ctr), Location(Ctr), HolderName(Ctr), ContactInfo(Ctr)
Loop
Close #1

End Sub

