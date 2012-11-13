VERSION 5.00
Begin VB.Form frmSearch
   BackColor       =   &H00404080&
   Caption         =   "Form1"
   ClientHeight    =   11175
   ClientLeft      =   60
   ClientTop       =   1950
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   11175
   ScaleWidth      =   19080
   Begin VB.CommandButton cmdClear
      Caption         =   "Clear"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   8
      Top             =   8400
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearchTitle
      Caption         =   "Search by Subject Title"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdSearchSource
      Caption         =   "Search by Source name"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1080
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   240
      Width           =   3135
   End
   Begin VB.PictureBox picContent
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   5760
      ScaleHeight     =   8835
      ScaleWidth      =   15555
      TabIndex        =   3
      Top             =   1920
      Width           =   15615
   End
   Begin VB.PictureBox picSubjectTitle
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   16320
      ScaleHeight     =   1515
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox picSourceName
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9360
      ScaleHeight     =   1515
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack
      Caption         =   "Back"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      TabIndex        =   0
      Top             =   9720
      Width           =   2535
   End
   Begin VB.Label lblSuggestTitle
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Subject Title: Guerre, Pretres, Jesuites, Style, Cacambo, Vieille"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   11
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label lblSuggestName
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SourceName: Candide, DP, JS, Pappas, Pearson"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   10
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label lblSuggestions
      BackColor       =   &H00FFFFC0&
      Caption         =   " Search Suggestions:"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   4335
   End
   Begin VB.Label lblSubjectTitle
      Caption         =   "Subject Title"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblSourceName
      Caption         =   "Source Name"
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NoteCard VB CS130 Project
'FrmSearch
'LauraBeach
'Written between February 1-24, 2010
'The purpose of the Search form is to read the data entered in the frmNewEntry and to allow the user to search that data to find where the data came from
'The form has Four buttons/commands. One returns the User to the frmMain (Main Page), One clears all of the picture boxes (where the results are displayed)
'The other two allow the User to search the data via Input Boxes by two separate variables...One searces by an the Subject array while the other searches the sources array



Option Explicit

Private Sub cmdBack_Click()

'Return to Main Form

frmMain.Show
frmSearch.Hide

End Sub

Private Sub cmdClear_Click()

'This Button Clears the contents of the picture boxes so that when another search is made it will be clear and more readable...

picSourceName.Cls
picSubjectTitle.Cls
picContent.Cls

End Sub

Private Sub cmdSearchSource_Click()
'Here it is necessary to dim a variety of variables as they are not public and need to be used just by this function...
    Dim CTRSource As Integer
    Dim Source As String
    Dim S As Integer
    Dim Found As Boolean

    Found = False

    CTRSource = 0
'here an InputBox function is used as the source of data for the variable aptly named 'Source'
    Source = InputBox("Please enter a Source to search for.")

'A For/Next Loop is used to perform an exhaustive search
    For S = 1 To CTR
        If Source = SourceName(S) Then
            CTRSource = CTRSource + 1
            Found = True
            picSourceName.Print CTRSource; ". "; UCase$(SourceName(S)) 'I have never used the String Functions so this is new for the UpperCase...
            picSubjectTitle.Print CTRSource; ". "; UCase$(Subject(S))
            picContent.Print CTRSource; ". "; Left(Content(S), 75); "..." 'The Left String Function helps here as the data is often larger/longer than the picture box
            'allows, giving a limit shows the page number (I asked that the page number be entered first in the data file, so that the page number might be seen as well as
            'the start of the quote...this makes the data 'findable' to the User who may need to refind it in the actual book/source of the data.)
        End If
    Next S
'This line of code provides for the fact that the notecard might not be found...
    If Found = False Then
        MsgBox "Sorry No Notecard was Found"
    End If




End Sub

Private Sub cmdSearchTitle_Click()

    Dim Found2 As Boolean
    Dim T As Integer
    Dim SubjectTitle As String
    Dim CTRSubject As Integer

    Found2 = False

    CTRSubject = 0

    SubjectTitle = InputBox("Please enter a Subject Title to search for.")

    For T = 1 To CTR
        If SubjectTitle = Subject(T) Then
            CTRSubject = CTRSubject + 1
            Found2 = True
            picSourceName.Print CTRSubject; ". "; UCase$(SourceName(T))
            picSubjectTitle.Print CTRSubject; ". "; UCase$(Subject(T))
            picContent.Print CTRSubject; ". "; Left(Content(T), 75); "..." 'The Left String Function helps here as the data is often larger/longer than the picture box
            'allows, giving a limit shows the page number (I asked that the page number be entered first in the data file, so that the page number might be seen as well as
            'the start of the quote...this makes the data 'findable' to the User who may need to refind it in the actual book/source of the data.)
        End If
    Next T
'This line of code provides for the fact that the notecard might not be found...
    If Found2 = False Then
        MsgBox "Sorry No Notecard was Found"
    End If

End Sub

Private Sub Label1_Click()

End Sub

