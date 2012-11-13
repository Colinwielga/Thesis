VERSION 5.00
Begin VB.Form frmTrafficProject 
   BackColor       =   &H80000011&
   Caption         =   "Traffic Project (1)"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   FillColor       =   &H00808080&
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   Picture         =   "VBTrafficProject.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   0
      Picture         =   "VBTrafficProject.frx":A606
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   14
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1695
      Left            =   9840
      Picture         =   "VBTrafficProject.frx":145AC
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   12
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1455
      Left            =   6000
      Picture         =   "VBTrafficProject.frx":20D72
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   3480
      Picture         =   "VBTrafficProject.frx":294D8
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   7080
      Width           =   2295
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   15480
      Picture         =   "VBTrafficProject.frx":345AA
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   6480
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   2280
      Picture         =   "VBTrafficProject.frx":3EBB0
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   5280
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   7560
      Picture         =   "VBTrafficProject.frx":47A72
      ScaleHeight     =   1035
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   7320
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000003&
      FillColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   4695
      Left            =   2160
      ScaleHeight     =   4635
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   240
      Width           =   12495
   End
   Begin VB.CommandButton cmdShowForm 
      Caption         =   "Show Form"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdAlphabetize 
      Caption         =   "Alphabetize"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearchPlate 
      Caption         =   "Search Plate"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpenSystem 
      Caption         =   "Open System"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
End
Attribute VB_Name = "frmTrafficProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Statement forces me to give meaning to every function in the program. 'Traffic Project Main Form 'Traffic Program: 'This program was written between the dates of March 16th, 2009 and March 24th, 2009 by Bill Roiger.
Dim CTR As Integer
Dim StudentInformation(1 To 5) As String, Title(1 To 5) As String, phonenumber(1 To 5) As String, studentidnumber(1 To 5) As String, homeaddress(1 To 5) As String, campusphone(1 To 5) As String, car(1 To 5) As String, licenseplate(1 To 5) As String 'The arrays.

Private Sub cmdAbout_Click() 'This command button brings up an explanation of what this program does, and why I chose to complete it.
    MsgBox "This is a simple representation of a vehicle/student program that is used at our campus as well as other campuses.  I chose to create this program because I'm employed with the CSB Security Department.", , "About The Program"
End Sub

Private Sub cmdAlphabetize_Click() 'This function alphabetizes all of the information in each of the arrays.
    Dim I As Integer, Pass As Integer, Pos As Integer
    Dim Temp As Integer, TempString As String
        For Pass = 1 To CTR - 1
            For Pos = 1 To (CTR - Pass) 'Bubble sort allows all array information to be sorted using temp storage areas.
                If Title(Pos) > Title(Pos + 1) Then
                    TempString = StudentInformation(Pos)
                    StudentInformation(Pos) = StudentInformation(Pos + 1)
                    StudentInformation(Pos + 1) = TempString
                    TempString = Title(Pos)
                    Title(Pos) = Title(Pos + 1)
                    Title(Pos + 1) = TempString
                    TempString = phonenumber(Pos)
                    phonenumber(Pos) = phonenumber(Pos + 1)
                    phonenumber(Pos + 1) = TempString
                    TempString = studentidnumber(Pos)
                    studentidnumber(Pos) = studentidnumber(Pos + 1)
                    studentidnumber(Pos + 1) = TempString
                    TempString = homeaddress(Pos)
                    homeaddress(Pos) = homeaddress(Pos + 1)
                    homeaddress(Pos + 1) = TempString
                    TempString = campusphone(Pos)
                    campusphone(Pos) = campusphone(Pos + 1)
                    campusphone(Pos + 1) = TempString
                    TempString = car(Pos)
                    car(Pos) = car(Pos + 1)
                    car(Pos + 1) = TempString
                    TempString = licenseplate(Pos)
                    licenseplate(Pos) = licenseplate(Pos + 1)
                    licenseplate(Pos + 1) = TempString
                End If
            Next Pos
        Next Pass
        picResults.Print "The alphabetized list is:"
        picResults.Print "name", , "phonenumber", "studentidnumber", "homeaddress", , "campusphone", "car", "licenseplate" 'Prints label headings for the outputs.
        picResults.Print "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" 'Underlines the label headings.
        For I = 1 To CTR
        picResults.Print Title(I); Tab(30); phonenumber(I); Tab(50); studentidnumber(I); Tab(70); homeaddress(I); Tab(110); campusphone(I); Tab(120); car(I); Tab(140); licenseplate(I) 'Neatly formats the outputs.
        Next I
End Sub

Private Sub cmdClear_Click() 'Clears everything in the picture box.
    picResults.Cls
End Sub

Private Sub cmdEnd_Click() 'Ends the program.
    End
End Sub

Private Sub cmdOpenSystem_Click()
Open App.Path & "\" & "StudentInformation.txt" For Input As #1 'Reads the student information into parallel arrays.
    CTR = 0
            picResults.Print "name", , "phonenumber", "studentidnumber", "homeaddress", , "campusphone", "car", "licenseplate" 'Prints label headings for the outputs.
            picResults.Print "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" 'Underlines the label headings.
        Do While Not EOF(1) 'Do while not loop, which ensures the array information will all be read into the system before the program quits reading.
            CTR = CTR + 1
            Input #1, Title(CTR), phonenumber(CTR), studentidnumber(CTR), homeaddress(CTR), campusphone(CTR), car(CTR), licenseplate(CTR)
            picResults.Print Title(CTR); Tab(30); phonenumber(CTR); Tab(50); studentidnumber(CTR); Tab(70); homeaddress(CTR); Tab(110); campusphone(CTR); Tab(120); car(CTR); Tab(140); licenseplate(CTR) 'Formats the print area.  'Tab functions ensure that the spacing will be neat and orderly.
        Loop
    Close #1 'Closes the application to the array information.
End Sub

Private Sub cmdSearchPlate_Click() 'Allows user to search for a specific license plate number in the system.
Dim Plate As String
Dim I As String
Dim Found As Boolean
Found = False
        Plate = InputBox("Enter a license plate number.", "License Plate") 'Inputbox function is where the user enters a plate number to search.
            I = 0
            Do While ((Not Found) And (I < CTR))
                I = I + 1
                If Plate = licenseplate(I) Then
                    Found = True
                End If
            Loop
picResults.Cls
                If (Not Found) Then 'If else statement is used to output a fitting result when a plate number is either found, or not found.
                    picResults.Print Plate; " is not in our system. "
                Else
                    picResults.Print Plate; " Belongs to "; Title(I)
                End If
End Sub

Private Sub cmdShowForm_Click() 'Hides the main form, and shows the sign-in name finder page.
frmTrafficProject.Hide
frmSignInName.Show
End Sub

