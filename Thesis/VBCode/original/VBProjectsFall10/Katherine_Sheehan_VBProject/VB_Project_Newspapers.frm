VERSION 5.00
Begin VB.Form frmReadFiles 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17115
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   17115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLastForm 
      Caption         =   "Return to the previous window"
      Height          =   735
      Left            =   960
      TabIndex        =   7
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdNextWindow 
      Caption         =   "Go to next window"
      Height          =   735
      Left            =   960
      TabIndex        =   6
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print selected information from all files"
      Height          =   735
      Left            =   960
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadArticleFile 
      Caption         =   "Read the file containing journal article sources"
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadBookFile 
      Caption         =   "Read the file containing book sources"
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadFile 
      Caption         =   "Read the file containing newspaper sources"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox picFileResults 
      Height          =   13215
      Left            =   4200
      ScaleHeight     =   13155
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   120
      Width           =   10455
   End
   Begin VB.Label lblFileInstructions 
      Caption         =   $"VB_Project_Newspapers.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmReadFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form reads all of the information into arrays and displays a selection of information from the sources

Private Sub cmdLastForm_Click()
    frmIntroduction.Show
    frmCommon.Hide
    frmReadFiles.Hide
    frmChicago.Hide
    'allows user to return to the previous form
End Sub

Private Sub cmdNextWindow_Click()
    frmReadFiles.Hide
    frmIntroduction.Hide
    frmCommon.Show
    frmChicago.Hide
    'shows the next form in the sequence and hides all the others
End Sub

Private Sub cmdPrint_Click()
    'this button prints selected results of each of the three files read
    'these results are meant to be a preview of sources available
    Dim NewsArticle As Single, JArticle As Single, Book As Single
    
    picFileResults.Print "You are using " & CTR & " newspaper sources, " & BookCTR & " book sources, and " & ArticleCTR & " journal article sources."
    picFileResults.Print " "
    picFileResults.Print "Author's Last Name"; Tab(25); "Year Published"; Tab(45); "Source Title"
    picFileResults.Print "---------------------------------------------------------------------------------------------------"
    'header for the chart
    For JArticle = 1 To ArticleCTR
        picFileResults.Print ArticleLN(JArticle); Tab(25); Year(JArticle); Tab(45); Article(JArticle)
    Next JArticle
    'prints selected results of journal article file
    
    For Book = 1 To BookCTR
        picFileResults.Print BookLastName(Book); Tab(25); YearPublished(Book); Tab(45); BookTitle(Book)
    Next Book
    'prints selected results of books file
    
    For NewsArticle = 1 To CTR
        picFileResults.Print LastName(NewsArticle); Tab(25); DatePublished(NewsArticle); Tab(45); ArticleTitle(NewsArticle)
    Next NewsArticle
    'prints selected results of newspaper file
End Sub

Private Sub cmdQuit_Click()
    End
    'ends the program
End Sub

Private Sub cmdReadArticleFile_Click()
    Open App.Path & "\Articles.txt" For Input As #3
    'opens articles file for input
    ArticleCTR = 0
    'sets counter equal to zero
    Do While Not EOF(3)
        ArticleCTR = ArticleCTR + 1
        'counts the number of articles used
        Input #3, Year(ArticleCTR), Journal(ArticleCTR), ArticleFN(ArticleCTR), ArticleLN(ArticleCTR), Volume(ArticleCTR), Article(ArticleCTR), PageNumbers(ArticleCTR), Issue(ArticleCTR)
        'saves information in a lot of arrays
    Loop
    Close #3
    'closes file
    MsgBox "The file has been read", , "File Read"
    'lets user know the file has been read successfully
End Sub

Private Sub cmdReadBookFile_Click()
    Dim Book As Single
    Open App.Path & "\Books.txt" For Input As #2
    'opens the file to be input
    BookCTR = 0
    'sets counter equal to zero
    Do While Not EOF(2)
        BookCTR = BookCTR + 1
        Input #2, BookTitle(BookCTR), YearPublished(BookCTR), PublishingTown(BookCTR), BookFirstName(BookCTR), BookLastName(BookCTR), PublishingCompany(BookCTR)
        'saves information from file in arrays
    Loop
    Close #2
    'closes the file so all of the information is saved
    MsgBox "The file has been read", , "File Read"
    'lets user know the file has been read successfully
End Sub

Private Sub cmdReadFile_Click()
    Open App.Path & "\Newspapers.txt" For Input As #1
    'opens files so information can be obtained
    CTR = 0
    'sets counter's initial value to zero
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, FirstName(CTR), LastName(CTR), DatePublished(CTR), ArticleTitle(CTR), Page(CTR), NewspaperName(CTR)
        'inputs and saves information from file
    Loop
    Close #1
    'closes the file because it is done obtaining information
    MsgBox "The file has been read", , "File Read"
End Sub
