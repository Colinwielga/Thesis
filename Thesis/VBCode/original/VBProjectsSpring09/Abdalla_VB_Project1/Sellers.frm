VERSION 5.00
Begin VB.Form frmSeller 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12075
   LinkTopic       =   "Form3"
   ScaleHeight     =   8790
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00404080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdRegister 
      BackColor       =   &H00000040&
      Caption         =   "Click here to start Registering the books you want to  sell"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   3975
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00004040&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   10155
      TabIndex        =   2
      Top             =   1920
      Width           =   10215
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   15
      Left            =   2760
      TabIndex        =   1
      Top             =   7320
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Register Books For Sell"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   5655
   End
End
Attribute VB_Name = "frmSeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Name: Book exchange'
'Form name: frmSeller'
'Author: Bibi Abdalla'
'Date: 3/24/2009'
'Objective: allow user to register thier books for sell'


Option Explicit
'This page allows seller to post thier books'
'seller state the name of the book'
'program checks to see if proffesor has registered the book by opening the booklist2.text'
'if the book is enlisted in booklist2.txt,the program asks the seller to state thier name, contact information and price they like to sell teh book for'
'the new data, plus the old data is saved in booklist3.text'
'if book is not enlisted, the program tell the sell that the book is not sellable. it then kicks the seller of the system'

'Declaring variables'

'varible for Entering books'
Dim Ctr As Integer
Dim Title As String
Dim Author As String
Dim Price As Double
Dim Field As String
Dim ProfName As String
Dim CourseName As String
Dim Location As String
Dim ISB As String
Dim HolderName As String
Dim Contact As String
'variable for Arrays'
Dim TitleArray(1 To 100) As String
Dim AuthorArray(1 To 100) As String
Dim PriceArray(1 To 10) As Double
Dim MarketPriceArray(1 To 10) As Double
Dim FieldArray(1 To 10) As String
Dim ProfNameArray(1 To 10) As String
Dim CourseNameArray(1 To 10) As String
Dim LocationArray(1 To 10) As String
Dim HolderNameArray(1 To 10) As String
Dim ContactInfoArray(1 To 10) As String
Dim ISBArray(1 To 100) As String
Dim Pass As Integer
Dim pos As Integer
Dim found As Boolean


'Quiting'

Private Sub cmdQuit_Click()
  frmSeller.Hide
  FrmWelcome.Show
  
End Sub

Private Sub cmdRegister_Click()
'Declaring variables'
Dim Index As Integer

'Opening File'
Open App.Path & "\booklist2.txt" For Input As #2
    Ctr = 0
'Writing data'
Do Until EOF(2)
    Ctr = Ctr + 1
    Input #2, TitleArray(Ctr), AuthorArray(Ctr), MarketPriceArray(Ctr), ISBArray(Ctr), FieldArray(Ctr), ProfNameArray(Ctr), CourseNameArray(Ctr), LocationArray(Ctr)
Loop
Close #2
'Operation'
'setting variables up'
    Title = InputBox("please enter title", "Tittle")
    Author = InputBox("Enter Author name", "Name of Author")
    found = False
    Pass = 0
    Index = 0
    'stating conditions'
    'comparing wheather data enter by user corresponds with data in booklist2.text'
    Do While ((Not found) And (Index < Ctr))
        Index = Index + 1
    If LCase(Title) = LCase(TitleArray(Index)) And LCase(Author) = LCase(AuthorArray(Index)) Then
        found = True
    'Printing data'
        picResult.Print "Title"; Tab(30); "Author"; Tab(60); "MarketPrice"; Tab(75); "Professor Name"; Tab(95); "CourseName"; Tab(120); "Location"
        'picResult.Print " Title ", " Author ", " Market Price", " ISB #", " Field ", " Professor ", " Course ", " Location"'
        picResult.Print "***************************************************************************************************************************************************************************************"
        'picResult.Print TitleArray(Index), AuthorArray(Index), FormatCurrency(MarketPriceArray(Index)), ISBArray(Index), FieldArray(Index), ProfNameArray(Index), , LocationArray(Index)'
        picResult.Print TitleArray(Index); Tab(30); AuthorArray(Index); Tab(60); FormatCurrency(MarketPriceArray(Index)); Tab(75); ProfNameArray(Index); Tab(95); CourseNameArray(Index); Tab(120); LocationArray(Index)
    End If
    Loop
    'setting conditions if statement is true'
    'if statement is true user is asked to enter name, price and contact information'
    If found = True Then
        HolderName = InputBox("Please enter the your name", "Your Name")
        Contact = InputBox("Please enter Your email", "Email")
        Price = InputBox("Please enter the price you like to sell the book for", "Price")
        'compare price with the market price stated in booklist2.text'
        found = False
        Do While Not found
           If Price < MarketPriceArray(Index) Then
                found = True
            Else
                Price = InputBox("your price is too expensive, Please enter another price", "Price")
            End If
        Loop
    'Storing data into booklist3.text'
    If found = True Then
        Open App.Path & "\booklist3.txt" For Append As #4
        Write #4, TitleArray(Index); AuthorArray(Index); MarketPriceArray(Index); Price; ISBArray(Index); FieldArray(Index); ProfNameArray(Index); CourseNameArray(Index); LocationArray(Index); HolderName; Contact
       
            'Printing data to allow user to see what they have entered'
           picResult.Print "_______________________________________________________________________________________________________________________________________________________________________________________________"
           picResult.Print "Title"; Tab(30); "Author"; Tab(60); "Price"; Tab(75); "MarketPrice"; Tab(95); "Professor Name"; Tab(120); "CourseName"; Tab(140)
           picResult.Print " ******************************************************************************************************************************************************************************************** "
           picResult.Print TitleArray(Index); Tab(30); AuthorArray(Index); Tab(60); FormatCurrency(Price); Tab(75); FormatCurrency(MarketPriceArray(Index)); Tab(95); ProfNameArray(Index); Tab(120); CourseNameArray(Index),
    End If
    End If
    'The statement,if book was not found in booklist2.text'
    If found = False Then
        MsgBox "Sorry this book is not sellable", , "Not Found"
    End If
   
End Sub
