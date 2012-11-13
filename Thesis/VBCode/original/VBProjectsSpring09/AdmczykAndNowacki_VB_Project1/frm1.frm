VERSION 5.00
Begin VB.Form frmUniversitySearch 
   BackColor       =   &H80000001&
   Caption         =   "Main Menu"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculator 
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Data "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort According to Tuition Price"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   8
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "Find Directions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton SearchPrivatePublic 
      Caption         =   "Search in Terms of Private or Public"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdMajor 
      Caption         =   "Search in Terms of Major"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      Height          =   3975
      Left            =   3720
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.CommandButton cmdSearchTuition 
      Caption         =   "Search in Terms of Tuition"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lbl2 
      BackColor       =   &H80000001&
      Caption         =   "Tuition:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lbl1 
      BackColor       =   &H80000001&
      Caption         =   "University:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmUniversitySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project name: College Bound
' Form name: University Search
' Authors: Magdalena Adamczyk & Leszek Nowacki
' 9-25 March 2009
' The main purpose of the program is to enable the user to search
' prospective universities around various states within the U.S.
' This form is designed to search through a list of universities according to the tuition price and major
' the user can also sort the results according to the tuition price
' it lets the user load other form on which he or she can search whether the school is public or private
' and load maps showing the universities

Option Explicit
Dim ctr As Single, I As Single, J As Single, TuitionSearched As Single, Searched As String
Dim pos As Single, pass As Single, temp As String, temp1 As Single, ctr2 As Single

Private Sub cmdCalculator_Click() ' this button hides the current form and displays other form which lets the user calculate the total costs of going to a university
frmUniversitySearch.Hide
frmCalculator.Show
End Sub

Private Sub cmdLoad_Click() ' this button loads the file with the list of universities in to arrays
Open App.Path & "\list.txt" For Input As #1 ' oppening the list of universities
Do Until EOF(1)     ' loading the file in to six parallel arrays
    ctr = ctr + 1
    Input #1, university(ctr), tuition(ctr), sort(ctr), major1(ctr), major2(ctr), major3(ctr)
Loop
Close #1 ' closing the file with the list of universities
cmdSearchTuition.Enabled = True ' enabling the button for searching the universities by the price of tuition
cmdMajor.Enabled = True 'enabling the button for searching the universities by major
End Sub

Private Sub cmdSearchTuition_Click() 'this button lets the user to search the universities for a major that he or she is interested in

picResults.Cls

TuitionSearched = InputBox("Please enter tuition you are willing to pay each semester", "Tuition") ' asking the user for major

Open App.Path & "\list2.txt" For Output As #2 'opening a file that will be used for storing the results onf the search


For I = 1 To ctr
    If tuition(I) <= TuitionSearched Then ' if tuition is smaler or equal to the tuition indicated by the user
    picResults.Print university(I); Tab(30); FormatCurrency(tuition(I)) 'the name of the university and the tuition that this university charges will be displayed
    Write #2, university(I), tuition(I), sort(I), major1(I), major2(I), major3(I) ' the universities that satisfy the conditions of the search will be saved in a file list2
    End If
Next I
Close #2 ' close the file list2

cmdSort.Enabled = True ' after a search is done and the universities that satisfy the conditions of the search are saved the button for sorting those results in enabled
End Sub

Private Sub cmdMajor_Click() ' this button lets the user search the list of the universities by major he or she wants to study

picResults.Cls

MajorSearched = InputBox("Please enter the major you want to study", "Major") ' a input box is asking the user to input desirable major

' as majors are listed in three arrays what makes it possible to asign three different majors to each university
' each of the arrays holding majors has to be checked it desirable major is present

For I = 1 To ctr
    If major1(I) = MajorSearched Then
    picResults.Print university(I); Tab(30); FormatCurrency(tuition(I))
    End If
Next I

For I = 1 To ctr
    If major2(I) = MajorSearched Then
    picResults.Print university(I); Tab(30); FormatCurrency(tuition(I))
    End If
Next I

For I = 1 To ctr
    If major3(I) = MajorSearched Then
    picResults.Print university(I); Tab(30); FormatCurrency(tuition(I))
    End If
Next I

Select Case MajorSearched ' for each possible major a different message is displayed as a coment to the user's choice
    Case Is = "art"
    MsgBox "Stick figures?--Come on!", , "Art"
    Case Is = "music"
    MsgBox "Can you really play something...besides video games?", , "Music"
    Case Is = "business"
    MsgBox "The name of the game is money!", , "Business"
    Case Is = "modern languages"
    MsgBox "Hola amigo!", , "Modern Languages"
    Case Is = "social sciences"
    MsgBox "Social Sciences?...Really?", , "Social Sciences"
    Case Is = "computer science"
    MsgBox "Loops, array and all that jazz!", , "Computer Science"
    Case Is = "communications"
    MsgBox "Hmm...taking it easy I see", , "Communications"
    Case Is = "law"
    MsgBox "No social life I guess", , "Law"
    Case Is = "medicine"
    MsgBox "A lot of all nighters", , "Medicine"
    Case Is = "environmental studies"
    MsgBox "Do you at least recicle?", , "Environmental Studies"
    Case Else
        MsgBox "There is no such major", , "Error!!!" ' if the major the user entered is not in any of the arrays holding majors an error massage is displayed
    End Select

Open App.Path & "\list2.txt" For Output As #2 'opening a file that will be used for storing the results onf the search

For I = 1 To ctr
    If major1(I) = MajorSearched Or major2(I) = MajorSearched Or major3(I) = MajorSearched Then ' if the desirable major is in any of the three arrays the university that offers this major
    Write #2, university(I), tuition(I), sort(I), major1(I), major2(I), major3(I) 'and other information associated with it will be saved to a file list2
    End If
Next I
Close #2

cmdSort.Enabled = True ' after a search is done and the universities that satisfy the conditions of the search are saved the button for sorting those results in enabled
End Sub

Private Sub cmdSort_Click() ' this button lets the user sort the result of his or hers search by the amount of tuition he or she wants to pay or by the major he or she wants to study
picResults.Cls


Open App.Path & "\list2.txt" For Input As #2 ' oppening the list of universities that satisfy the conditions of the user's search
ctr2 = 0
Do Until EOF(2)
    ctr2 = ctr2 + 1
    Input #2, university2(ctr2), tuition2(ctr2), sort2(ctr2), major12(ctr2), major22(ctr2), major32(ctr2) ' loading the file in to six parallel arrays
Loop
Close #2 ' closing file list2


For pass = 1 To ctr2 - 1 ' sorting the arrays according to tuition starting with the lowest tuition at the beggining
    For pos = 1 To ctr2 - pass
        If tuition2(pos) > tuition2(pos + 1) Then
            temp = university2(pos)     ' two temp variables are used one as string for the university's name, its kind and majors
            university2(pos) = university2(pos + 1)
            university2(pos + 1) = temp
            temp1 = tuition2(pos)       'the second temp which is a single is used for storing the amount of tuition
            tuition2(pos) = tuition2(pos + 1)
            tuition2(pos + 1) = temp1
            temp = sort2(pos)
            sort2(pos) = sort2(pos + 1)
            sort2(pos + 1) = temp
            temp = major12(pos)
            major12(pos) = major12(pos + 1)
            major12(pos + 1) = temp
            temp = major12(pos)
            major22(pos) = major22(pos + 1)
            major22(pos + 1) = temp
            temp = major32(pos)
            major32(pos) = major32(pos + 1)
            major32(pos + 1) = temp
        End If
    Next pos
Next pass


For J = 1 To ctr2
    picResults.Print university2(J); Tab(30); FormatCurrency(tuition2(J)) 'the universities and amount of tuition in a sorted version is then displayed
Next J
End Sub

Private Sub SearchPrivatePublic_Click() 'this button hides the current form and displays other form which makes possible finding a university according to its kind
frmUniversitySearch.Hide
frmPrivateOrPublic.Show
End Sub

Private Sub cmdMap_Click() ' this button hides the current form and displays other form which lets the user see maps showing the location of different universities
frmUniversitySearch.Hide
frmDirections.Show

End Sub

Private Sub cmdQuit_Click() 'quit button
End
End Sub

