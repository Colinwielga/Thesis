VERSION 5.00
Begin VB.Form People 
   BackColor       =   &H00FFC0FF&
   Caption         =   "People"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form2"
   ScaleHeight     =   8745
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdshow 
      BackColor       =   &H0080FFFF&
      Caption         =   "A Final Picture of Our Group"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H0080FFFF&
      Caption         =   "Find a Friend"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdlstname 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Friends by Last Name"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdrmnumber 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Friends by Room Number"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdgender 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Friends by Gender"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdpeople 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Friends Alphabetically by First Name"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0080FFFF&
      Caption         =   "Back to Home Page"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFC0FF&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8355
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Ashley 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ashley K. Smithson"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "People"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Australia
'Form Name: People
'Author: Ashley Smithson
'Date: October 31, 2005
'Purpose of form: To show the students I traveled with and be able to find their room number or name by alphabetical order first and last name.
Option Explicit
Dim pass As Single, CTR As Integer, COmp As Single, list As Single, A As Integer
Dim names(1 To 40) As String, gender(1 To 40) As String, roomnumber(1 To 40) As Integer, tempgender As String, tempname As String, temproomnumber As Integer
Dim lstnames(1 To 40) As String, templstnames As String, X As String, placeCTR As Integer, Found As Boolean

Private Sub cmdback_Click()
People.Hide 'brings the user back to the home page
FinalProject2.Show
End Sub

Private Sub cmdfind_Click()
X = InputBox("Enter the last name of who you are looking for:", "Last Name")
Open App.Path & "\people.txt" For Input As #1
'takes the name from the user, searches the file, people for the information
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, names(CTR), lstnames(CTR), gender(CTR), roomnumber(CTR)
Loop
Close #1 'closes the file
picresults.Cls 'clears the picresults box
picresults.Print "First Name", "Last Name", "Gender", "Room Number"
picresults.Print "**********************************************************************"
'prints the information on the top as a header
Found = False
placeCTR = 0
Do While (Not Found) And placeCTR < CTR
placeCTR = placeCTR + 1
    If lstnames(placeCTR) = X Then
    picresults.Print names(placeCTR), lstnames(placeCTR), gender(placeCTR), roomnumber(placeCTR)
    Found = True
End If
Loop
'searches for the name entered if found prints resulting info
If Not Found Then
    picresults.Print "No", X, "here"
End If
'searches for name entered if not found shows they are not there
End Sub

Private Sub cmdgender_Click()
Open App.Path & "\people.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, names(CTR), lstnames(CTR), gender(CTR), roomnumber(CTR)
Loop
'counter
Close #1
picresults.Cls
picresults.Print "First Name", "Last Name", "Gender", "Room Number"
picresults.Print "**********************************************************************"
For pass = 1 To CTR - 1
    For COmp = 1 To CTR - pass
        If gender(COmp) > gender(COmp + 1) Then
            tempgender = gender(COmp)
            gender(COmp) = gender(COmp + 1)
            gender(COmp + 1) = tempgender
            tempname = names(COmp)
            names(COmp) = names(COmp + 1)
            names(COmp + 1) = tempname
            temproomnumber = roomnumber(COmp)
            roomnumber(COmp) = roomnumber(COmp + 1)
            roomnumber(COmp + 1) = temproomnumber
            templstnames = lstnames(COmp)
            lstnames(COmp) = lstnames(COmp + 1)
            lstnames(COmp + 1) = templstnames
        End If
    Next COmp
Next pass
'puts in order... males and females
For A = 1 To CTR
  picresults.Print names(A), lstnames(A), gender(A), roomnumber(A)
Next A
'prints information in order
End Sub



Private Sub cmdlstname_Click()
Open App.Path & "\people.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, names(CTR), lstnames(CTR), gender(CTR), roomnumber(CTR)
Loop
Close #1
picresults.Cls
picresults.Print "First Name", "Last Name", "Gender", "Room Number"
picresults.Print "**********************************************************************"
For pass = 1 To CTR - 1
    For COmp = 1 To CTR - pass
        If lstnames(COmp) > lstnames(COmp + 1) Then
            templstnames = lstnames(COmp)
            lstnames(COmp) = lstnames(COmp + 1)
            lstnames(COmp + 1) = templstnames
            tempname = names(COmp)
            names(COmp) = names(COmp + 1)
            names(COmp + 1) = tempname
            tempgender = gender(COmp)
            gender(COmp) = gender(COmp + 1)
            gender(COmp + 1) = tempgender
            temproomnumber = roomnumber(COmp)
            roomnumber(COmp) = roomnumber(COmp + 1)
            roomnumber(COmp + 1) = temproomnumber
            
        End If
    Next COmp
Next pass
'put in order by last name
For A = 1 To CTR
  picresults.Print names(A), lstnames(A), gender(A), roomnumber(A)
Next A
End Sub

Private Sub cmdpeople_Click()
Open App.Path & "\people.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, names(CTR), lstnames(CTR), gender(CTR), roomnumber(CTR)
Loop
Close #1
picresults.Cls
picresults.Print "First Name", "Last Name", "Gender", "Room Number"
picresults.Print "**********************************************************************"
For pass = 1 To CTR - 1
    For COmp = 1 To CTR - pass
        If names(COmp) > names(COmp + 1) Then
            tempname = names(COmp)
            names(COmp) = names(COmp + 1)
            names(COmp + 1) = tempname
            tempgender = gender(COmp)
            gender(COmp) = gender(COmp + 1)
            gender(COmp + 1) = tempgender
            temproomnumber = roomnumber(COmp)
            roomnumber(COmp) = roomnumber(COmp + 1)
            roomnumber(COmp + 1) = temproomnumber
            templstnames = lstnames(COmp)
            lstnames(COmp) = lstnames(COmp + 1)
            lstnames(COmp + 1) = templstnames
        End If
    Next COmp
Next pass
'read in information from notebook file and put in order by first name
For A = 1 To CTR
  picresults.Print names(A), lstnames(A), gender(A), roomnumber(A)
Next A
    
End Sub

Private Sub cmdrmnumber_Click()
Open App.Path & "\people.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, names(CTR), lstnames(CTR), gender(CTR), roomnumber(CTR)
Loop
Close #1
picresults.Cls
picresults.Print "First Name", "Last Name", "Gender", "Room Number"
picresults.Print "**********************************************************************"
For pass = 1 To CTR - 1
    For COmp = 1 To CTR - pass
        If roomnumber(COmp) > roomnumber(COmp + 1) Then
            temproomnumber = roomnumber(COmp)
            roomnumber(COmp) = roomnumber(COmp + 1)
            roomnumber(COmp + 1) = temproomnumber
            tempname = names(COmp)
            names(COmp) = names(COmp + 1)
            names(COmp + 1) = tempname
            tempgender = gender(COmp)
            gender(COmp) = gender(COmp + 1)
            gender(COmp + 1) = tempgender
            templstnames = lstnames(COmp)
            lstnames(COmp) = lstnames(COmp + 1)
            lstnames(COmp + 1) = templstnames
        End If
    Next COmp
Next pass
'read in information and put in order by room number
For A = 1 To CTR
  picresults.Print names(A), lstnames(A), gender(A), roomnumber(A)
Next A
End Sub

Private Sub cmdshow_Click()
picresults.Picture = LoadPicture(App.Path & "\Pictures\FinalGroup.jpg")
'load and display group picture
End Sub

