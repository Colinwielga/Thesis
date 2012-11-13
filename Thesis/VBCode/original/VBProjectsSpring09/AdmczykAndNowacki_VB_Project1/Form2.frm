VERSION 5.00
Begin VB.Form frmPrivateOrPublic 
   BackColor       =   &H00079304&
   Caption         =   "Private or Public University"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort in Terms of Tuition"
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
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "Get Directions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
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
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack1 
      Caption         =   "<< Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      Height          =   3135
      Left            =   3600
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000000&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H80000000&
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox cbo1 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "Private/Public"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00079304&
      Caption         =   "Tuition"
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
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00079304&
      Caption         =   "University"
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
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00079304&
      Caption         =   "What type of University would you like?"
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
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmPrivateOrPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project name: College Bound
' Form name: University Search
' Authors: Magdalena Adamczyk & Leszek Nowacki
' 9-25 March 2009
' this form makes possible for the user to find a only private universities or only public universities

Option Explicit
Dim ctr As Single, tuition(1 To 10) As Single, I As Single, sortSearched As String
Dim pos As Single, pass As Single, temp As String, temp1 As Single, ctr2 As Single, J As Single
Private Sub Form_Load() ' at the time of loading this form the program is designed to load the content of the scroll down list(combobox)
cbo1.AddItem "private"
cbo1.AddItem "public"

Open App.Path & "\list.txt" For Input As #1 ' when the form is loaded also the file with the list of the universities is loaded into parallel arrays
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, university(ctr), tuition(ctr), sort(ctr), major1(ctr), major2(ctr), major3(ctr)
Loop
Close #1

End Sub

Private Sub cbo1_Click() ' when user choices one of the positions from the scroll down list the sort button becomes disenabled until the user accepts his choice by clicking OK
cmdSort.Enabled = False
End Sub

Private Sub cmdOK_Click()

picResults.Cls
sortSearched = cbo1.Text ' the type of the university that the user searches is read from the combo box


Open App.Path & "\list2.txt" For Output As #2 'opening a file that will be used for storing the results onf the search

For I = 1 To ctr
    If sort(I) = sortSearched Then
        Write #2, university(I), tuition(I), sort(I), major1(I), major2(I), major3(I) ' the universities that are of the appropriate type the conditions of the search will be saved in a file list2
        picResults.Print university(I); Tab(30); FormatCurrency(tuition(I)) 'the name of the university and the tuition that this university charges will be displayed
    End If
Next I
Close #2 ' closing the file list2

cmdSort.Enabled = True
End Sub

Private Sub cmdSort_Click()
picResults.Cls

Open App.Path & "\list2.txt" For Input As #2 'opening a file that will be used for storing the results onf the search
ctr2 = 0
Do Until EOF(2)
    ctr2 = ctr2 + 1
    Input #2, university2(ctr2), tuition2(ctr2), sort2(ctr2), major12(ctr2), major22(ctr2), major32(ctr2) ' loading the file in to six parallel arrays
Loop
Close #2 ' closing the file list2


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
    picResults.Print university2(J); Tab(30); FormatCurrency(tuition2(J))
Next J

End Sub

Private Sub cmdMap_Click() ' this button hides the current form and displays other form which lets the user see maps showing the location of different universities
frmDirections.Show
frmPrivateOrPublic.Hide
End Sub

Private Sub cmdBack1_Click() ' this button hides the current form and displays other form which lets the user search through a list of universities according to the tuition price and major
frmUniversitySearch.Show
frmPrivateOrPublic.Hide
End Sub

Private Sub cmdQuit_Click() 'this button ends the program
End
End Sub

