VERSION 5.00
Begin VB.Form NewSnowForm 
   Caption         =   "Wax Project by John Boruff"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "NewSnowFrom.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAlready 
      BackColor       =   &H0000FFFF&
      Caption         =   "If you already have wax, see what conditions it is good for"
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdChart 
      BackColor       =   &H0000FFFF&
      Caption         =   "View Wax Chat"
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdWax 
      BackColor       =   &H0000FFFF&
      Caption         =   "Find the right wax for the tempature"
      Height          =   1575
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox picResultswax 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   3840
      ScaleHeight     =   5715
      ScaleWidth      =   5235
      TabIndex        =   6
      Top             =   0
      Width           =   5295
   End
   Begin VB.PictureBox picTemp 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   6960
      Width           =   6015
   End
   Begin VB.TextBox txtTemp 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Text            =   "32"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdReturnprevious 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to previous screen"
      Height          =   1095
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturnmain 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to main screen"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdTemp 
      BackColor       =   &H0000FFFF&
      Caption         =   "Convert"
      Height          =   1095
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Not Shure what the tempature is in celsius.Enter a tempature in Farnheit to Conver to centigrade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2655
   End
End
Attribute VB_Name = "NewSnowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : WaxProject (John Boruff's VB-project.vbp)
'Form Name : NewSnowForm (NewSnowForm.frm)
'Author: John Boruff
'purpose of the form:  Since the user has selected a snow type, the purpose
                    'of this form is for the user to determine which type of wax
                    'he/she needs depending on the tempature in degrees celsius.

Dim Celsius(1 To 23) As Single, Fast(1 To 23) As String, Medium(1 To 23) As String, Slow(1 To 23) As String, I As Integer, Temp As Single, Wax As String, Found As Boolean
Public Path As String

Private Sub cmdAlready_Click()
picResultswax.Cls   'Clears any previous text in the print box
Wax = InputBox("Please enter the type of wax you have, refer to wax guide to ensure proper name", "New and Fine Snow") 'gets the wax type from the user
Found = False 'sets found equal to false
picResultswax.Print "You can use "; Wax; " For speed at all costs when the air tempature is"
For I = 1 To 23
    If Wax = Left(Fast(I), 5) Then  'Searchs array to see if the wac can be used as a speed at all costs wax'
        Found = True  'sets equation to true if wax type is found'
        picResultswax.Print Celsius(I)  'prints the tempature at which the wax type was found at
    End If
Next I
For I = 1 To 23  'begins the loop
    If Wax = Left(Fast(I), 4) And (Len(Fast(I)) = 18 Or Len(Fast(I)) = 19) Then  'Searchs array to see if the wac can be used as a speed at all costs wax'
        Found = True  'sets equation to true if wax type is found'
        picResultswax.Print Celsius(I)  'prints the tempature at which the wax type was found at
    End If
Next I
For I = 1 To 23
    If Wax = Right(Fast(I), 8) And Len(Fast(I)) = 18 Then  'Searchs array to see if the wac can be used as a speed at all costs wax'
        Found = True  'sets equation to true if wax type is found'
        picResultswax.Print Celsius(I)  'prints the tempature at which the wax type was found at
    End If
Next I
For I = 1 To 23
    If Wax = Right(Fast(I), 9) And Len(Fast(I)) = 19 Then  'Searchs array to see if the wac can be used as a speed at all costs wax'
        Found = True  'sets equation to true if wax type is found'
        picResultswax.Print Celsius(I) 'prints the tempature at which the wax type was found at
    End If
Next I
For I = 1 To 23
    If Wax = Right(Fast(I), 9) And Len(Fast(I)) = 20 Then  'Searchs array to see if the wac can be used as a speed at all costs wax'
        Found = True  'sets equation to true if wax type is found'
        picResultswax.Print Celsius(I) 'prints the tempature at which the wax type was found at
    End If
Next I
If Found = False Then
    picResultswax.Print "I'm sorry, "; Wax; " is not a speed at all costs wax" 'if the wax type wasn't found prints this statement
End If
picResultswax.Print ""  'if the wax type was found prints the following statement
picResultswax.Print "If you want a medium wax you can use "; Wax; " in these tempatures"
Found = False  'sets found back to fas=lse incase it was changed to true
For I = 1 To 23
    If Wax = Left(Medium(I), 5) Then  'searches Medium array to see if the wax type could be used as a medium wax
        Found = True  'If wax type is found to true
        picResultswax.Print Celsius(I)  'prints tempature at which wax could be used at
    End If
Next I
If Found = False Then
   picResultswax.Print "I'm sorry, "; Wax; " is not a medium wax"  'prints message if found equals false
End If
picResultswax.Print ""
picResultswax.Print "If you want a inexpensive wax you can use "; Wax; " in these tempatures"
Found = False 'If wax type is found changes found back to false
For I = 1 To 23
    If Wax = Left(Slow(I), 5) Then  'searches slow array for the wax type
        Found = True  'If wax type is found to true
        picResultswax.Print Celsius(I)  'prints the tempatures at which the wax type was found at
    End If
Next I
If Found = False Then
   picResultswax.Print "I'm sorry,"; Wax; " is not a inexpensive wax" 'If no wax was found prints this statement
End If
picResultswax.Print ""
picResultswax.Print "***** If the temapture 6 appeared you should"  'prints a message to help user understand the waxing instructions
picResultswax.Print "use the wax for 6 degrees or or higher tempatures"
picResultswax.Print " and if -16 appeares use the wax for -16 degrees or lower."
    
End Sub
Private Sub cmdChart_Click()
picResultswax.Cls  'Clears any previous text is picture box
I = 1
picResultswax.Print "New and Fine Snow Wax Chart"
picResultswax.Print "Temp"; Tab(10); "Speed at all costs"; Tab(34); "Medium"; Tab(45); "Inexpensive"
Do While Not I = 24  'prints out array into table format so user can see the wax chart
    picResultswax.Print Celsius(I); Tab(10); Fast(I); Tab(34); Medium(I); Tab(45); Slow(I)
    I = I + 1
Loop
picResultswax.Print ""
picResultswax.Print "***** If the temapture 6 appeared you should" 'prints a message to help user understand the waxing instructions
picResultswax.Print "use the wax for 6 degrees or or higher tempatures"
picResultswax.Print " and if -16 appeares use the wax for -16 degrees or lower."
End Sub

Private Sub cmdReturnmain_Click()
    MainForm1.Show  'returns user to the main form
    NewSnowForm.Hide
End Sub

Private Sub cmdReturnprevious_Click()
    SnowForm.Show  'returns user to the SnowForm
    NewSnowForm.Hide
End Sub

Private Sub cmdTemp_Click()
Dim F As Single, C As Single
    'converts fahrnheit into Celsius
    picTemp.Cls
    F = txtTemp
    C = (5 / 9) * (F - 32)
    picTemp.Print F; " degrees fahrnheit ="; FormatNumber(C, 2); "degrees centigrade"
End Sub

Private Sub cmdWax_Click()
picResultswax.Cls
Temp = InputBox("Enter the air tempature in celsius", "New And Fine Snow")
picResultswax.Print "Temp"; Tab(10); "Speed At All Cost"; Tab(34); "Medium"; Tab(45); "Economical"
Select Case Temp  'user enters a tempature and comapres it if statements untilit fits an equatuion than prints wax types for that tempature
    Case Is <= -13
        picResultswax.Print Temp; Tab(10); Fast(1); Tab(34); Medium(1); Tab(45); Slow(1)
    Case Is <= -11
        picResultswax.Print Temp; Tab(10); Fast(6); Tab(34); Medium(6); Tab(45); Slow(6)
    Case Is <= -8
        picResultswax.Print Temp; Tab(10); Fast(9); Tab(34); Medium(9); Tab(45); Slow(9)
    Case Is <= -6
        picResultswax.Print Temp; Tab(10); Fast(11); Tab(34); Medium(11); Tab(45); Slow(11)
    Case Is < -2
        picResultswax.Print Temp; Tab(10); Fast(15); Tab(34); Medium(15); Tab(45); Slow(15)
    Case Is <= -1
        picResultswax.Print Temp; Tab(10); Fast(16); Tab(34); Medium(16); Tab(45); Slow(16)
    Case Is <= 3
        picResultswax.Print Temp; Tab(10); Fast(20); Tab(34); Medium(20); Tab(45); Slow(20)
    Case Is <= 5
        picResultswax.Print Temp; Tab(10); Fast(22); Tab(34); Medium(22); Tab(45); Slow(22)
    Case Is > 5
        picResultswax.Print Temp; Tab(10); Fast(23); Tab(34); Medium(23); Tab(45); Slow(23)
End Select
If Temp >= -12 Then
    picResultswax.Print ""  'prints a message to help user understand the waxing instructions
    picResultswax.Print "Cerra F should only be applied  in situations"
    picResultswax.Print "when speed is is crucial, such as race day because"
    picResultswax.Print "of its extreme cost. Using the first wax listed under"
    picResultswax.Print "speed at all costs and not applying Cerra F should give"
    picResultswax.Print "most everyone the speed they are looking for."
    End If
End Sub

Private Sub Form_Load()  'opens file so arrays can be read in
Path = "M:\CS130\Boruff_John\"
Open App.Path & "\NewandFine.txt" For Input As #1
For I = 1 To 23
    Input #1, Celsius(I), Fast(I), Medium(I), Slow(I)
Next I
Close #1  'closes file when done reading in array
End Sub
