VERSION 5.00
Begin VB.Form frmChicago 
   BackColor       =   &H00400040&
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoForward 
      Caption         =   "Go to the next window"
      Height          =   975
      Left            =   840
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdGoAuthor 
      Caption         =   "Go"
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdGoTitle 
      Caption         =   "Go"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to previous window"
      Height          =   975
      Left            =   3960
      TabIndex        =   8
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtAuthor 
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtTitle 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picFinalResults 
      Height          =   6615
      Left            =   7680
      ScaleHeight     =   6555
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   360
      Width           =   6735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Quit"
      Height          =   975
      Left            =   3960
      TabIndex        =   1
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Enter the last name of an author to see Chicago style in-text citations for all of their work"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblEntireTitle 
      Caption         =   "Enter the entire title of a source to get a Chicago style in-text citation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblFormInstructions 
      Caption         =   $"frmChicago.frx":0000
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
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "frmChicago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form creates Chicago style in-text citations after the user enters either the entire title of the source or
'the author's last name

Private Sub cmdBack_Click()
    frmChicago.Hide
    frmCommon.Show
    frmIntroduction.Hide
    frmReadFiles.Hide
    frmBibliography.Hide
    'allows user to return to the previous window
End Sub

Private Sub cmdClear_Click()
    picFinalResults.Cls
    txtTitle.Text = " "
    txtAuthor.Text = " "
    'clears all text boxes and the picture box
End Sub

Private Sub cmdEnd_Click()
    End
    'ends the program
End Sub

Private Sub cmdGoAuthor_Click()
    Dim Author As String, Found As Boolean, N As Single, A As Single, B As Single
    'variables only to be used for this button
    Author = txtAuthor.Text
    'variable is set to the information entered by the user
    
    picFinalResults.Print " "
    picFinalResults.Print "Sources written by '" & Author & "':"
    picFinalResults.Print " "
    'prints header for the information
    
    For N = 1 To CTR
        If LCase(Author) = LCase(LastName(N)) Then
            Found = True
            picFinalResults.Print FirstName(N) & " " & LastName(N) & ", '" & ArticleTitle(N) & ", ' " & NewspaperName(N) & ", " & DatePublished(N) & ", " & Page(N)
        End If
    Next N
    'searches for the last name in the newspaper arrays
    For A = 1 To ArticleCTR
        If LCase(Author) = LCase(ArticleLN(A)) Then
            Found = True
            picFinalResults.Print ArticleFN(A) & " " & ArticleLN(A) & ", '" & Article(A) & ",' " & Journal(A) & ", " & Volume(A) & ", " & Issue(A) & " (" & Year(A) & "):" & PageNumbers(A)
        End If
    Next A
    'searches for last name in the article group
    For B = 1 To BookCTR
        If LCase(Author) = LCase(BookLastName(B)) Then
            Found = True
            picFinalResults.Print BookFirstName(B) & " " & BookLastName(B) & ", " & BookTitle(B) & " (" & PublishingTown(B) & ": " & PublishingCompany(B) & " " & YearPublished(B) & "),"
        End If
    Next B
    'searches book arrays for the last name
    
    If (Not Found) Then
        MsgBox Author & " was not found. Please enter a different last name.", , "Not Found"
    End If
    'tells the program what to do if the author is not found
End Sub

Private Sub cmdGoForward_Click()
    frmReadFiles.Hide
    frmIntroduction.Hide
    frmCommon.Hide
    frmChicago.Hide
    frmBibliography.Show
    'user moves on to next window
End Sub

Private Sub cmdGoTitle_Click()
    Dim Title As String, Found As Boolean, N As Single, A As Single, B As Single
    'variables to be used for this button only
    Title = txtTitle.Text
    'obtains the information the user wrote in the text box
    
    picFinalResults.Print " "
    picFinalResults.Print "Sources with the title '" & Title & "':"
    picFinalResults.Print " "
    'prints header for the information
    
    Found = False
    'sets found value to false to start off with so match/stop search can be used
    N = 0
    A = 0
    B = 0
    'all values must be set to zero because 3 different match/stop searches are going to be used
    'because there were 3 different input files
    
    Do While ((Not Found) And (N < CTR))
        N = N + 1
        If LCase(Title) = LCase(ArticleTitle(N)) Then
            Found = True
            picFinalResults.Print FirstName(N) & " " & LastName(N) & ", '" & ArticleTitle(N) & ", ' " & NewspaperName(N) & ", " & DatePublished(N) & ", " & Page(N)
        End If
    Loop
    'searches newspaper arrays for title
    Do While ((Not Found) And (A < ArticleCTR))
        A = A + 1
        If LCase(Title) = LCase(Article(A)) Then
            Found = True
            picFinalResults.Print ArticleFN(A) & " " & ArticleLN(A) & ", '" & Article(A) & ",' " & Journal(A) & ", " & Volume(A) & ", " & Issue(A) & " (" & Year(A) & "):" & PageNumbers(A)
        End If
    Loop
    'searches journal articles for title
    Do While ((Not Found) And (B < BookCTR))
        B = B + 1
        If LCase(Title) = LCase(BookTitle(B)) Then
            Found = True
            picFinalResults.Print BookFirstName(B) & " " & BookLastName(B) & " '" & BookTitle(B) & "' (" & PublishingTown(B) & ": " & PublishingCompany(B) & " " & YearPublished(B) & "),"
        End If
    Loop
    'searches books for title entered
    
    If (Not Found) Then
        MsgBox Title & " was not found. Please enter a different title.", , "Title Not Found"
    End If
End Sub

Private Sub cmdTest_Click()
    Dim A As Single
    For A = 1 To ArticleCTR
        picFinalResults.Print ArticleFN(A), ArticleLN(A), Journal(A), Volume(A), Issue(A), Year(A)
    Next A
    
End Sub
