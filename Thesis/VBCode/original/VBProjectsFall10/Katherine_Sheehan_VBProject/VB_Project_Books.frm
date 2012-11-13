VERSION 5.00
Begin VB.Form frmCommon 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   14550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreviousWindow 
      Caption         =   "Return to the previous window"
      Height          =   855
      Left            =   3720
      TabIndex        =   14
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdNextForm 
      Caption         =   "Go to the next window"
      Height          =   855
      Left            =   1320
      TabIndex        =   13
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoYear 
      Caption         =   "Go"
      Height          =   615
      Left            =   5040
      TabIndex        =   12
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdGoPublisher 
      Caption         =   "Go"
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdGoLastName 
      Caption         =   "Go"
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   3720
      TabIndex        =   9
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdClearPic 
      Caption         =   "Clear"
      Height          =   855
      Left            =   1320
      TabIndex        =   8
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtYear 
      Height          =   735
      Left            =   2760
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox txtPublisher 
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   6375
      Left            =   6360
      ScaleHeight     =   6315
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   480
      Width           =   7335
   End
   Begin VB.TextBox txtLastName 
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblCommonInstructions 
      Caption         =   $"VB_Project_Books.frx":0000
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
      Left            =   1080
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblYear 
      Caption         =   "Enter the year you would like sources from (because all newspapers are from the same year, they are not included in this search):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblPublisher 
      Caption         =   "Enter the publisher, journal, or newspaper you would like to see work from:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblLastName 
      Caption         =   "Enter the last name of the author you would like to see work from:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form searches the arrays for matches to the publisher, year, or author entered by the user.

Private Sub cmdClearPic_Click()
    picResults.Cls
    'clears the picture box
    txtLastName.Text = " "
    txtPublisher.Text = " "
    txtYear.Text = " "
    'clears all of the text boxes all well as the picture box
End Sub
Private Sub cmdGoLastName_Click()
    Dim LN As String, Found As Boolean, News As Single, LnCTR As Single, Art As Single, B As Single
    'i will be searching each of the three files for the last name of the author
    'variables will only be used in this form

    LN = txtLastName.Text
    'information for LastName will be obtained from text box
    picResults.Print "Sources written by " & LN & ":"
    picResults.Print " "
    'header for the information
    LnCTR = 0
    'sets intitial counter value to zero
    For News = 1 To CTR
        If (LCase(LN) = LCase(LastName(News))) Then
            Found = True
            LnCTR = LnCTR + 1
            picResults.Print FirstName(News) & " " & LN & " has written '" & ArticleTitle(News) & "'"
        End If
    Next News
    'searches for the last name in the newspaper arrays
    For Art = 1 To ArticleCTR
        If (LCase(LN) = LCase(ArticleLN(Art))) Then
            Found = True
            LnCTR = LnCTR + 1
            picResults.Print ArticleFN(Art) & " " & LN & " has written '" & Article(Art) & "'"
        End If
    Next Art
    'searches for last name in the article group
    For B = 1 To BookCTR
        If (LCase(LN) = LCase(BookLastName(B))) Then
            Found = True
            LnCTR = LnCTR + 1
            picResults.Print BookFirstName(B) & " " & LN & " has written '" & BookTitle(B) & "'"
        End If
    Next B
    'searches book arrays for the last name
    
    If (Not Found) Then
        MsgBox LN & " was not found. Please enter a different name.", , "Error"
    End If
    'shows the user an error message if the last name is not in the information
End Sub

Private Sub cmdGoPublisher_Click()
    'this button deals with the publisher, journal, and newspaper because it is searching for who the work was published by
    'this makes all of the information appropriate for a single button
    Dim Publisher As String, A As Single, B As Single, N As Single, Found As Boolean, PCTR As Single
    'variables will only be used for this button
    'once again, searching through information from all three files
    
    Publisher = txtPublisher.Text
    'loops will use information gathered from this text box
    picResults.Print "Sources published by " & Publisher & ":"
    picResults.Print " "
    'prints header for the information
    
    PCTR = 0
    'sets the initial counter equal to zero
    
    For A = 1 To ArticleCTR
        If (LCase(Publisher) = LCase(Journal(A))) Then
            Found = True
            PCTR = PCTR + 1
            picResults.Print Publisher & " published '" & Article(A) & "' by " & ArticleFN(A) & " " & ArticleLN(A) & " in " & Year(A)
        End If
    Next A
    'searches information from the article file
    For B = 1 To BookCTR
        If (LCase(Publisher) = LCase(PublishingCompany(B))) Then
            Found = True
            PCTR = PCTR + 1
            picResults.Print Publisher & " published '" & BookTitle(B) & "' by " & BookFirstName(B) & " " & BookLastName(B) & " in " & YearPublished(B)
        End If
    Next B
    'searches for publisher in the book arrays
    For N = 1 To CTR
        If (LCase(Publisher) = LCase(NewspaperName(N))) Then
            Found = True
            PCTR = PCTR + 1
            picResults.Print Publisher & " published '" & ArticleTitle(N) & "' by " & FirstName(N) & " " & LastName(N) & " on "; DatePublished(N)
        End If
    Next N
    'searches through newspaper information for the publisher
    
    If (Not Found) Then
        MsgBox Publisher & " was not found. Please enter another publisher, journal, or newspaper.", , "Error"
    End If
    'tells the program what to do if the user inputs information not contained within the arrays
End Sub

Private Sub cmdGoYear_Click()
    Dim YR As Single, A As Single, B As Single, YrCTR As Single, Found As Boolean
    'these variables will only be used by this button
    'i used the same variables for the for-next loop for continuity and also to help myself remember
    'which variables go with which file
    YR = txtYear.Text
    'YR is set to whatever is written in the year text box
    picResults.Print "Sources published in " & YR & ":"
    picResults.Print " "
    'header for the information to be printed
    
    YrCTR = 0
    'counter's initial value is zero
    
    For A = 1 To ArticleCTR
        If (YR = Year(A)) Then
            Found = True
            YrCTR = YrCTR + 1
            picResults.Print "'" & Article(A) & "' by " & ArticleFN(A) & " " & ArticleLN(A) & " was published in " & YR
        End If
    Next A
    'searches through journal articles
    For B = 1 To BookCTR
        If (YR = YearPublished(B)) Then
            Found = True
            YrCTR = YrCTR + 1
            picResults.Print "'" & BookTitle(B) & "' by " & BookFirstName(B) & " " & BookLastName(B); " was published in " & YR
        End If
    Next B
    'searches books for the year entered in the text box
    
    If (Not Found) Then
        MsgBox YR & " was not found. Please enter a different year.", , "Error"
    End If
    'show user this message if the year was not entered in the information from the files
End Sub

Private Sub cmdNextForm_Click()
    frmChicago.Show
    frmCommon.Hide
    frmIntroduction.Hide
    frmReadFiles.Hide
    frmBibliography.Hide
    'users proceeds to the next window
End Sub

Private Sub cmdPreviousWindow_Click()
    frmIntroduction.Hide
    frmCommon.Hide
    frmReadFiles.Show
    frmChicago.Hide
    frmBibliography.Hide
    'allows user to go back to the previous window
End Sub

Private Sub cmdQuit_Click()
    End
    'ends the program from this window
End Sub

