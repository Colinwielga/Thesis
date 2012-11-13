VERSION 5.00
Begin VB.Form frmAboutMPCC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Minnesota Private College Council, Fund, and Research Foundation"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLocation 
      BackColor       =   &H000000C0&
      Caption         =   "Show Schools by Location in MN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdEnrollment 
      BackColor       =   &H000000C0&
      Caption         =   "Show Schools by 08 Undergraduate Enrollment"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdtuition 
      BackColor       =   &H000000C0&
      Caption         =   "Show Schools by 2008-2009 Tuition"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H000000C0&
      Caption         =   "Show Schools from A to Z"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin VB.PictureBox picSlideshow 
      Height          =   1575
      Left            =   2160
      Picture         =   "AboutMPCC.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   2355
      TabIndex        =   9
      Top             =   4920
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4335
      Left            =   7320
      ScaleHeight     =   4275
      ScaleWidth      =   6195
      TabIndex        =   7
      Top             =   4560
      Width           =   6255
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H000000C0&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H000000C0&
      Caption         =   "Want to know more specifically about one of our schools? Click here!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   3015
   End
   Begin VB.CommandButton cmdResearch 
      BackColor       =   &H000000C0&
      Caption         =   "Information for Policymakers and Researchers"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton cmdEmployment 
      BackColor       =   &H000000C0&
      Caption         =   "Interested in Employment? "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click one of the options below to learn more about all of our member schools:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   10
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label lblLink 
      BackStyle       =   0  'Transparent
      Caption         =   "www.mnprivatecolleges.org"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   8760
      Width           =   3495
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"AboutMPCC.frx":1554
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   6375
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About the Minnesota Private College Council, Fund, and Research Foundation:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   9135
   End
End
Attribute VB_Name = "frmAboutMPCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SchoolName(1 To 20) As String
Dim enrollment(1 To 20) As Double
Dim tuition(1 To 20) As Double
Dim location(1 To 20) As String
Dim ctr As Integer

'   Day at the Capitol and MN Private College Information Tool
'   Form: AboutMPCC
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective: The objective of this form is to provide other information routes about the MN Private College Council, including
'   information about each school's enrollment, and tools to do further research, find employment, and register for the MPCC's
'   Day at the Capitol for each private college.



Private Sub cmdAlpha_Click()
picResults.Cls 'Clears results if other sorts have been completed in the picture box.

'this subroutine sorts the names alphabetically and
'maintains the correct information in the arrays for each school.
Dim pass As Integer, pos As Integer, j As Integer
Dim tempSchoolName As String, tempenrollment As Integer, temptuition As Double, templocation As String

'sort the names
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If SchoolName(pos) > SchoolName(pos + 1) Then
            tempSchoolName = SchoolName(pos)
            SchoolName(pos) = SchoolName(pos + 1)
            SchoolName(pos + 1) = tempSchoolName
            tempenrollment = enrollment(pos)
            enrollment(pos) = enrollment(pos + 1)
            enrollment(pos + 1) = tempenrollment
            temptuition = tuition(pos)
            tuition(pos) = tuition(pos + 1)
            tuition(pos + 1) = temptuition
            templocation = location(pos)
            location(pos) = location(pos + 1)
            location(pos + 1) = templocation
        End If
    Next pos
Next pass
 
'print the list
'first print the header info
    picResults.Print "School"; Tab(38); "Enrollment"; Tab(55); "Tuition"; Tab(70); "Location"; Tab(85)
    picResults.Print "*************************************************************************************************"
    
'then print the list
    For j = 1 To ctr
             picResults.Print SchoolName(j); Tab(38); enrollment(j); Tab(55); FormatCurrency(tuition(j)); Tab(70); location(j); Tab(85)
    Next j

End Sub

Private Sub cmdEmployment_Click()
frmAboutMPCC.Hide
frmEmployment.Show

End Sub

Private Sub cmdEnrollment_Click()
Dim pass As Integer, pos As Integer, j As Integer
Dim tempSchoolName As String, tempenrollment As Integer, temptuition As Double, templocation As String

picResults.Cls 'Clears results if other sorts have been completed in the picture box.

'This subroutine sorts the enrollment in ascending order and
'maintains the correct information in the parallel arrays for each school.

'sort the names
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If enrollment(pos) > enrollment(pos + 1) Then
            tempenrollment = enrollment(pos)
            enrollment(pos) = enrollment(pos + 1)
            enrollment(pos + 1) = tempenrollment
            tempSchoolName = SchoolName(pos)
            SchoolName(pos) = SchoolName(pos + 1)
            SchoolName(pos + 1) = tempSchoolName
            temptuition = tuition(pos)
            tuition(pos) = tuition(pos + 1)
            tuition(pos + 1) = temptuition
            templocation = location(pos)
            location(pos) = location(pos + 1)
            location(pos + 1) = templocation
        End If
    Next pos
Next pass
 
'print the list
'first print the header info
    picResults.Print "School"; Tab(38); "Enrollment"; Tab(55); "Tuition"; Tab(70); "Location"; Tab(85)
    picResults.Print "*************************************************************************************************"
    
'then print the list
    For j = 1 To ctr
             picResults.Print SchoolName(j); Tab(38); enrollment(j); Tab(55); FormatCurrency(tuition(j)); Tab(70); location(j); Tab(85)
    Next j


End Sub

Private Sub cmdFind_Click()
Dim School As String

Dim user As String
School = InputBox("Please enter the full name of the institution you would like to see.", "What school would you like more information about?")
    If School = "Augsburg College" Then
        frmAboutMPCC.Hide
        frmAugsburg.Show
    ElseIf School = "Bethany Lutheran College" Then
        frmAboutMPCC.Hide
        frmBethany.Show
    ElseIf School = "Bethel University" Then
        frmAboutMPCC.Hide
        frmBethel.Show
    ElseIf School = "Carleton College" Then
        frmAboutMPCC.Hide
        frmCarleton.Show
    ElseIf School = "College of St. Benedict" Then
        frmAboutMPCC.Hide
        frmCSBSJU.Show
    ElseIf School = "College of St. Catherine" Then
        frmAboutMPCC.Hide
        frmStCat.Show
    ElseIf School = "College of St. Scholastica" Then
        frmAboutMPCC.Hide
        frmScholastica.Show
    ElseIf School = "Concordia College, Moorhead" Then
        frmAboutMPCC.Hide
        frmMoorhead.Show
    ElseIf School = "Concordia University, St. Paul" Then
        frmAboutMPCC.Hide
        frmStPaul.Show
    ElseIf School = "Gustavus Adolphus College" Then
        frmAboutMPCC.Hide
        frmGAC.Show
    ElseIf School = "Hamline University" Then
        frmAboutMPCC.Hide
        frmHamline.Show
    ElseIf School = "Macalester College" Then
        frmAboutMPCC.Hide
        frmMacalester.Show
    ElseIf School = "Minneapolis College of Art and Design" Then
        frmAboutMPCC.Hide
        frmMCAD.Show
    ElseIf School = "St. John's University" Then
        frmAboutMPCC.Hide
        frmCSBSJU.Show
    ElseIf School = "St. Mary's University" Then
        frmAboutMPCC.Hide
        frmStMary.Show
    ElseIf School = "St. Olaf College" Then
        frmHome.Hide
        frmStOlaf.Show
    ElseIf School = "University of St. Thomas" Then
        frmAboutMPCC.Hide
        frmStThomas.Show
    Else
        MsgBox "Sorry! Please enter the full name of the institution.", , "Error!"
        
    End If


End Sub

Private Sub txtInfo_Change()
txtInfo = "About us: The Minnesota Private College Council, Fund and Research Foundation are related non-profit organizations that represent private higher education in Minnesota. Members of the organizations are 17 private, four-year liberal arts colleges. Our mission is to preserve and enhance quality private higher education to serve the education and economic needs of our region. The organizations share a common goal: to create policy and funding conditions which allow any qualified Minnesota student the opportunity to attend a Minnesota private college. We serve our member colleges through research, policy advocacy, positioning efforts and facilitation of institution collaboration — including through raising funds to support our students and institutions."

End Sub

Private Sub cmdHome_Click()
frmAboutMPCC.Hide
frmHome.Show

End Sub

Private Sub cmdLocation_Click()
Dim pass As Integer, pos As Integer, j As Integer
Dim tempSchoolName As String, tempenrollment As Integer, temptuition As Double, templocation As String

picResults.Cls 'Clears results if other sorts have been completed in the picture box.

'This subroutine sorts the schools by location in MN (city name) and
'maintains the correct information in the arrays for each school.

'sort the names
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If location(pos) > location(pos + 1) Then
            templocation = location(pos)
            location(pos) = location(pos + 1)
            location(pos + 1) = templocation
            tempenrollment = enrollment(pos)
            enrollment(pos) = enrollment(pos + 1)
            enrollment(pos + 1) = tempenrollment
            tempSchoolName = SchoolName(pos)
            SchoolName(pos) = SchoolName(pos + 1)
            SchoolName(pos + 1) = tempSchoolName
            temptuition = tuition(pos)
            tuition(pos) = tuition(pos + 1)
            tuition(pos + 1) = temptuition
        End If
    Next pos
Next pass
 
'print the list
'first print the header info
    picResults.Print "School"; Tab(38); "Enrollment"; Tab(55); "Tuition"; Tab(70); "Location"; Tab(85)
    picResults.Print "*************************************************************************************************"
    
'then print the list
    For j = 1 To ctr
             picResults.Print SchoolName(j); Tab(38); enrollment(j); Tab(55); FormatCurrency(tuition(j)); Tab(70); location(j); Tab(85)
    Next j


End Sub

Private Sub cmdResearch_Click()

frmAboutMPCC.Hide
frmResearch.Show

End Sub

Private Sub cmdtuition_Click()
Dim pass As Integer, pos As Integer, j As Integer
Dim tempSchoolName As String, tempenrollment As Integer, temptuition As Double, templocation As String


picResults.Cls 'Clears results if other sorts have been completed in the picture box.

'This subroutine sorts the schools by tuition cost and
'maintains the correct information in the arrays for each school.

'sort the names
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If tuition(pos) > tuition(pos + 1) Then
            temptuition = tuition(pos)
            tuition(pos) = tuition(pos + 1)
            tuition(pos + 1) = temptuition
            tempenrollment = enrollment(pos)
            enrollment(pos) = enrollment(pos + 1)
            enrollment(pos + 1) = tempenrollment
            tempSchoolName = SchoolName(pos)
            SchoolName(pos) = SchoolName(pos + 1)
            SchoolName(pos + 1) = tempSchoolName
            templocation = location(pos)
            location(pos) = location(pos + 1)
            location(pos + 1) = templocation
        End If
    Next pos
Next pass
 
'print the list
'first print the header info
    picResults.Print "School"; Tab(38); "Enrollment"; Tab(55); "Tuition"; Tab(70); "Location"; Tab(85)
    picResults.Print "*************************************************************************************************"
    
'then print the list
    For j = 1 To ctr
             picResults.Print SchoolName(j); Tab(38); enrollment(j); Tab(55); FormatCurrency(tuition(j)); Tab(70); location(j); Tab(85)
    Next j


End Sub

Private Sub Form_Load()

'Open the file of information and put it in arrays called SchoolNames, enrollment, tuition, and location.
Open App.Path & "\Enrollment.txt" For Input As #1

ctr = 0

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, SchoolName(ctr), enrollment(ctr), tuition(ctr), location(ctr)
Loop
Close #1


End Sub

Private Sub lblAbout_Click()

End Sub

'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Private Sub lblLink_Click()
Const url As String = "http://www.mnprivatecolleges.org"

    Set ie = CreateObject("internetexplorer.application")
    With ie
        .Visible = True
        .navigate url
    End With
    Set ie = Nothing


End Sub


Private Sub Quit_Click()
End
End Sub
