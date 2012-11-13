VERSION 5.00
Begin VB.Form GlossaryANDMeter 
   BackColor       =   &H00FF8080&
   Caption         =   "GlossaryANDMeter"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Instructions 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4560
      ScaleHeight     =   1515
      ScaleWidth      =   5115
      TabIndex        =   35
      Top             =   2400
      Width           =   5175
   End
   Begin VB.CommandButton TeacherInst 
      Caption         =   "Get Instructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      TabIndex        =   34
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   31
      Top             =   6600
      Width           =   1095
   End
   Begin VB.PictureBox GBox 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3960
      ScaleHeight     =   1515
      ScaleWidth      =   5715
      TabIndex        =   30
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   29
      Top             =   6600
      Width           =   1095
   End
   Begin VB.PictureBox MeterResult 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      ScaleHeight     =   1155
      ScaleWidth      =   4995
      TabIndex        =   28
      Top             =   5280
      Width           =   5055
   End
   Begin VB.CommandButton MeterTest 
      Caption         =   "Is My Measure Right?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   27
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox R1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   19
      Top             =   6240
      Width           =   615
   End
   Begin VB.TextBox R2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   18
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox R4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   17
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox R8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   16
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox N1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   15
      Top             =   6240
      Width           =   615
   End
   Begin VB.TextBox R16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   14
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox N2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox N4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   12
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox N8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox N16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Top             =   4320
      Width           =   630
   End
   Begin VB.OptionButton MeterIn4 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   255
   End
   Begin VB.OptionButton MeterIn2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   4320
      Width           =   255
   End
   Begin VB.OptionButton MeterIn3 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Dictionary 
      Caption         =   "Look Up Term"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Signature 
      BackStyle       =   0  'Transparent
      Caption         =   "Created By: Thomas Jones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9840
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label GlossaryLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Musical Term Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Whole(Dotted Half In 3/4)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   26
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Half"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Eighth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   24
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sixteenth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Quarter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   22
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label NoteLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The Number of Each Note or Rest"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label TimeLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Time Signature"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "4/4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "3/4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "2/4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label MeterLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2520
      Width           =   2775
   End
End
Attribute VB_Name = "GlossaryANDMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Classroom Music Aide
'Meter And Glossary
'Thomas R. Jones
'Written:10/29/03
'The purpose of this porgram is to be aid a classrom music teacher by allowing students to have an interactive model from which to learn various musical concepts.
'It is the intent of the author that this form will eventually serve as but a part of a more comprehensive musical program.

Option Explicit 'Forces the programmer to define all variables so that no assumptions are made by the program itself.
Dim Path As String
Dim Final As Integer
Dim WholeNote As Integer
Dim HalfRest As Integer
Dim SixteenthNote As Integer
Dim SixteenthRest As Integer
Dim EighthNote As Integer
Dim EighthRest As Integer
Dim QuarterNote As Integer
Dim QuarterRest As Integer
Dim HalfNote As Integer
Dim WholeRest As Integer
Private Sub Clear_Click() 'Clears both the glossary and meter result boxes.
MeterResult.Cls
GBox.Cls
Instructions.Cls
End Sub

Private Sub Form_Load() 'variable path used for grading purposes
Path = "M:\CS130\Thomas Jones\"
End Sub

Private Sub MeterTest_Click()
MeterResult.Cls
If MeterIn2.Value = True Then 'Checks to see if the 2/4 meter option has been chosen
    WholeNote = Val(N1.Text)
    If WholeNote > 0 Then 'Determines if the user has incorrectly used a whole note to represent an entire measure in 2/4
        MeterResult.Print "We Use Half Notes To Show Full Measures In 2/4"
    End If 'Ends the whole note check
    HalfRest = Val(R2.Text)
    If HalfRest > 0 Then 'Checks to see if half rest has been used incorrectly in 2/4
        MeterResult.Print "We Use A Quarter Rest To Show Half A Measure Of Rest In 2/4"
    End If 'Ends half rest check
    SixteenthNote = Val(N16.Text) 'The following set all variable names to the corresponding text box values
    SixteenthRest = Val(R16.Text)
    EighthNote = Val(N8.Text)
    EighthRest = Val(R8.Text)
    QuarterNote = Val(N4.Text)
    QuarterRest = Val(R4.Text)
    HalfNote = Val(N2.Text)
    WholeRest = Val(R1.Text)
    'The following sets use the basis that there are 8 sixteenth notes in a full measure of 2/4.
    'Values are added then the result is printed
    Final = (SixteenthNote * 1) + (SixteenthRest * 1) + (EighthNote * 2) + (EighthRest * 2) + (QuarterNote * 4) + (QuarterRest * 4) + (HalfNote * 8) + (WholeRest * 8)
    If Final = 8 Then
        MeterResult.Print "YES!" 'Prints results to meter box
    Else
        MeterResult.Print "NO!"
    End If
End If
    
If MeterIn3.Value = True Then 'Checks to see that 3/4 option has been selected
        SixteenthNote = Val(N16.Text) 'Again, variables are set to text box values
        SixteenthRest = Val(R16.Text)
        EighthNote = Val(N8.Text)
        EighthRest = Val(R8.Text)
        QuarterNote = Val(N4.Text)
        QuarterRest = Val(R4.Text)
        HalfNote = Val(N2.Text)
        HalfRest = Val(R2.Text)
        WholeNote = Val(N1.Text)
        WholeRest = Val(R1.Text)
    'This time the program uses the basis of 12 sixteenth notes in a total bar
    'The values of the selected notes and rests are added are result is determined
    Final = (SixteenthNote * 1) + (SixteenthRest * 1) + (EighthNote * 2) + (EighthRest * 2) + (QuarterNote * 4) + (QuarterRest * 4) + (HalfNote * 8) + (HalfRest * 8) + (WholeNote * 12) + (WholeRest * 12)
    If Final = 12 Then
        MeterResult.Print "YES!" 'Again prints results to the meter box
    Else
        MeterResult.Print "NO!"
    End If
End If
If MeterIn4.Value = True Then 'Checks to see that the 4/4 option has been selected
        SixteenthNote = Val(N16.Text) 'Again, variables are set to text box values
        SixteenthRest = Val(R16.Text)
        EighthNote = Val(N8.Text)
        EighthRest = Val(R8.Text)
        QuarterNote = Val(N4.Text)
        QuarterRest = Val(R4.Text)
        HalfNote = Val(N2.Text)
        HalfRest = Val(R2.Text)
        WholeNote = Val(N1.Text)
        WholeRest = Val(R1.Text)
    'Here, the program assumes a full bar has 16 total sixteenth note counts
    'The values of the selected notes and rests are again added and the result is determined
    Final = (SixteenthNote * 1) + (SixteenthRest * 1) + (EighthNote * 2) + (EighthRest * 2) + (QuarterNote * 4) + (QuarterRest * 4) + (HalfNote * 8) + (HalfRest * 8) + (WholeNote * 16) + (WholeRest * 16)
    If Final = 16 Then
        MeterResult.Print "YES!" 'Again, results are printed to the meter box
    Else
        MeterResult.Print "NO!"
    End If 'Ends result and print if statment
End If 'Ends first if statment
End Sub
Private Sub Dictionary_Click()
Dim NotFound As Boolean 'Notice that the Term is set as a string and the other I as a counter
Dim T As String
Dim I As Double
Dim Term(1 To 40) As String 'Defines Term Array
Dim Definition(1 To 40) As String 'Defines Definition Array
'The next section opens the text file and puts it into two parallel arrays
Open Path & "wordlist.txt" For Input As #1
For I = 1 To 40
    Input #1, Term(I), Definition(I)
    Next I
Close #1
T = InputBox("Please Enter A Musical Term(Use All Lower Case)", "Term") 'Asks User For A Term
'Begins a sequential search
I = 0
NotFound = True
Do While NotFound And I < 40 'Searches until term is found or end of list is reached
    I = I + 1
    If T = Term(I) Then NotFound = False
Loop
If NotFound Then 'Prints the results
    GBox.Cls
        GBox.Print T; "This Term Is Not Currently Included In The Glossary"
    Else
    GBox.Cls
        GBox.Print T, ; Definition(I);
    End If
End Sub

Private Sub Quit_Click()
End
End Sub

Private Sub TeacherInst_Click() 'Gives preset and changable instructions ot the student
Instructions.Print "1. Create a measure in each time signature"
Instructions.Print "                                                "
Instructions.Print "2. Use Quarter, Eighth, and Sixteenth notes/rests in your examples"
Instructions.Print "                                                "
Instructions.Print "3. Make Up Some Of Your Own"
End Sub
