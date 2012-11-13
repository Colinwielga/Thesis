VERSION 5.00
Begin VB.Form frmBibliography 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinalReturn 
      Caption         =   "Return to the previous window"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Program"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write bibliography to a text file"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblFinalInstructions 
      Caption         =   $"frmBibliography.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   5775
   End
End
Attribute VB_Name = "frmBibliography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form writes a bibliography to a text file, and lets the user end the program. it is the last form.

Private Sub cmdFinalReturn_Click()
    frmReadFiles.Hide
    frmIntroduction.Hide
    frmCommon.Hide
    frmChicago.Show
    frmBibliography.Hide
    'lets user return to previous window
End Sub

Private Sub cmdQuit_Click()
    End
    'ends the program
End Sub

Private Sub cmdWrite_Click()
    Dim N As Single, A As Single, B As Single
    Dim TempBLN As String, TempBFN As String, TempBTitle As String, TempTown As String, TempCompany As String, TempBYear As Single
    Dim Pass As Single, Pos As Single, TempLN As String, TempFN As String, TempArticleTitle As String, TempNewspaper As String, TempDate As String
    Dim TempALN As String, TempAFN As String, TempArticle As String, TempJournal As String
    Dim TempVolume As Single, TempIssue As String, TempYear As Single
    'variables to be used for this button only
    For Pass = 1 To BookCTR - 1
        For Pos = 1 To BookCTR - Pass
            If BookLastName(Pos) > BookLastName(Pos) Then
                TempBLN = BookLastName(Pos)
                BookLastName(Pos) = BookLastName(Pos + 1)
                BookLastName(Pos + 1) = TempBLN
                
                TempBFN = BookFirstName(Pos)
                BookFirstName(Pos) = BookFirstName(Pos + 1)
                BookFirstName(Pos + 1) = TempBFN
                
                TempBTitle = BookTitle(Pos)
                BookTitle(Pos) = BookTitle(Pos + 1)
                BookTitle(Pos + 1) = TempBTitle
                
                TempTown = PublishingTown(Pos)
                PublishingTown(Pos) = PublishingTown(Pos + 1)
                PublishingTown(Pos + 1) = TempTown
                
                TempCompany = PublishingCompany(Pos)
                PublishingCompany(Pos) = PublishingCompany(Pos + 1)
                PublishingCompany(Pos + 1) = TempCompany
                
                TempBYear = YearPublished(Pos)
                YearPublished(Pos) = YearPublished(Pos + 1)
                YearPublished(Pos + 1) = TempBYear
            End If
        Next Pos
    Next Pass
    'sorts books into alphabetical order
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If LastName(Pos) > LastName(Pos + 1) Then
                TempLN = LastName(Pos)
                LastName(Pos) = LastName(Pos + 1)
                LastName(Pos + 1) = TempLN
                
                TempFN = FirstName(Pos)
                FirstName(Pos) = FirstName(Pos + 1)
                FirstName(Pos + 1) = TempFN
                
                TempArticleTitle = ArticleTitle(Pos)
                ArticleTitle(Pos) = ArticleTitle(Pos + 1)
                ArticleTitle(Pos + 1) = TempArticleTitle
                
                TempNewspaper = NewspaperName(Pos)
                NewspaperName(Pos) = NewspaperName(Pos + 1)
                NewspaperName(Pos + 1) = TempNewspaper
                
                TempDate = DatePublished(Pos)
                DatePublished(Pos) = DatePublished(Pos + 1)
                DatePublished(Pos + 1) = TempDate
            End If
        Next Pos
    Next Pass
    'sorts newpapers into alphabetical order
    For Pass = 1 To ArticleCTR - 1
        For Pos = 1 To ArticleCTR - Pass
            If ArticleLN(Pos) > ArticleLN(Pos + 1) Then
                TempALN = ArticleLN(Pos)
                ArticleLN(Pos) = ArticleLN(Pos + 1)
                ArticleLN(Pos + 1) = TempALN
                    
                TempAFN = ArticleFN(Pos)
                ArticleFN(Pos) = ArticleFN(Pos + 1)
                ArticleFN(Pos + 1) = TempAFN
                    
                TempArticle = Article(Pos)
                Article(Pos) = Article(Pos + 1)
                Article(Pos + 1) = TempArticle
                    
                TempJournal = Journal(Pos)
                Journal(Pos) = Journal(Pos + 1)
                Journal(Pos + 1) = TempJournal
                    
                TempVolume = Volume(Pos)
                Volume(Pos) = Volume(Pos + 1)
                Volume(Pos + 1) = TempVolume
                    
                TempIssue = Issue(Pos)
                Issue(Pos) = Issue(Pos + 1)
                Issue(Pos + 1) = TempIssue
                    
                TempYear = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = TempYear
            End If
        Next Pos
    Next Pass
    'sorts articles into alphabetical order
    Open App.Path & "\Bibliography.txt" For Output As #4
    'opens file to be written to
    For N = 1 To CTR
        Write #4, LastName(N), FirstName(N), ArticleTitle(N), NewspaperName(N), DatePublished(N)
    Next N
    'sorts the newpspaer articles in alphabetical order and writes them to the file
    
    For A = 1 To ArticleCTR
        Write #4, ArticleLN(A), ArticleFN(A), Article(A), Journal(A), Volume(A), Issue(A), Year(A)
    Next A
    ' and writes article info to the file

    For B = 1 To BookCTR
        Write #4, BookLastName(B), BookFirstName(B), BookTitle(B), PublishingTown(B), PublishingCompany(B), YearPublished(B)
    Next B
    'writes book info to the file
    Close #4
    MsgBox "A bibligraphy has been written to the file 'Bibliography.txt'", , "Bibliography Written"
End Sub
