VERSION 5.00
Begin VB.Form OldSnowForm 
   Caption         =   "Wax Project by John Boruff"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "OldSnowForm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReturnprevious2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Previous Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9600
      Width           =   2655
   End
   Begin VB.PictureBox picTemp 
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
      Height          =   1215
      Left            =   3000
      ScaleHeight     =   1155
      ScaleWidth      =   3915
      TabIndex        =   9
      Top             =   7800
      Width           =   3975
   End
   Begin VB.CommandButton cmdReturnprevious 
      Caption         =   "Return to previous screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   15600
      TabIndex        =   8
      Top             =   9840
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturnmain 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to main screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9600
      Width           =   2655
   End
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H0000FFFF&
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   2535
   End
   Begin VB.TextBox txtTemp 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Text            =   "32"
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdWax 
      BackColor       =   &H0000FFFF&
      Caption         =   "Find the right wax for the temperature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdChart 
      BackColor       =   &H0000FFFF&
      Caption         =   "View Wax Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdAlready 
      BackColor       =   &H0000FFFF&
      Caption         =   "If you already have wax, see what conditions it's good for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   3975
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
      Height          =   6855
      Left            =   9480
      ScaleHeight     =   6795
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label lblTemp 
      BackColor       =   &H0000FFFF&
      Caption         =   "Not sure what the temperature is in celsius?  Enter the temperature in Farnheit to convert to centigrade."
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   3255
   End
End
Attribute VB_Name = "OldSnowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : WaxProject (John Boruff's VB-project.vbp)
'Form Name : OldSnowForm (OldSnowForm.frm)
'Author: John Boruff
'purpose of the form:  Since the user has selected a snow type, the purpose
                    'of this form is for the user to determine which type of wax
                    'he/she needs depending on the tempature in degrees celsius.

Dim Celsius(1 To 23) As Single, Fast(1 To 23) As String, Medium(1 To 23) As String, Slow(1 To 23) As String, I As Integer, Temp As Single, Wax As String, Found As Boolean, Path As String
Private Sub cmdAlready_Click()
picResultswax.Cls  'clears any previous text from pic box
Wax = InputBox("Please enter the type of wax you have, refer to wax guide to ensure proper name", "Old, Slush, and Icy Snow")
Found = False  'sets equation equal to false
picResultswax.Print "You can use "; Wax; " For speed at all costs when the air tempature is"
For I = 1 To 23
    If Wax = Left(Fast(I), 5) Then 'searches fast array for the wax type
        Found = True  'if wax type is found than changes found to true
        picResultswax.Print Celsius(I) 'prints out the temp at which wax type was found
    End If
Next I
For I = 1 To 23
    If Wax = Left(Fast(I), 4) And (Len(Fast(I)) = 18 Or Len(Fast(I)) = 19) Then
        Found = True  'if wax type is found than changes found to true
        picResultswax.Print Celsius(I)  'prints out the temp at which wax type was found
    End If
Next I
For I = 1 To 23
    If Wax = Right(Fast(I), 8) And Len(Fast(I)) = 18 Then
        Found = True  'if wax type is found than changes found to true
        picResultswax.Print Celsius(I) 'prints out the temp at which wax type was found
    End If
Next I
For I = 1 To 23
    If Wax = Right(Fast(I), 9) And Len(Fast(I)) = 19 Then
        Found = True  'if wax type is found than changes found to true
        picResultswax.Print Celsius(I) 'prints out the temp at which wax type was found
    End If
Next I
For I = 1 To 23
    If Wax = Right(Fast(I), 9) And Len(Fast(I)) = 20 Then
        Found = True   'if wax type is found than changes found to true
        picResultswax.Print Celsius(I) 'prints out the temp at which wax type was found
    End If
Next I
If Found = False Then
    picResultswax.Print "I'm sorry, "; Wax; " is not a speed at all costs wax" 'if found = false prints statement
End If
picResultswax.Print ""
picResultswax.Print "If you want a medium wax you can use "; Wax; " in these tempatures"
Found = False  'sets fouond beck to false if it was changes to true
For I = 1 To 23
    If Wax = Left(Medium(I), 5) Then  'searches medium array for the specified wax type
        Found = True 'if wax type is found than changes found to true
        picResultswax.Print Celsius(I) 'prints out the temp at which wax type was found
    End If
Next I
If Found = False Then 'prints following is wax waxn't found
   picResultswax.Print "I'm sorry, "; Wax; " is not a medium wax"
End If
picResultswax.Print ""  'prints following if wax was found
picResultswax.Print "If you want a inexpensive wax you can use "; Wax; " in these tempatures"
Found = False  'sets found back to false if it was changes to true
For I = 1 To 23
    If Wax = Left(Slow(I), 5) Then  'searches slow array for wax type
        Found = True 'if wax type is found than changes found to true
        picResultswax.Print Celsius(I) 'prints out the temp at which wax type was found
    End If
Next I
If Found = False Then  'prints following if wax type wasn't found
   picResultswax.Print "I'm sorry, "; Wax; " is not a inexpensive wax"
End If
picResultswax.Print ""  'prints a message to help user understand the waxing instructions
picResultswax.Print "***** If the temapture 6 appeared you should"
picResultswax.Print "use the wax for 6 degrees or or higher tempatures"
picResultswax.Print " and if -16 appeares use the wax for -16 degrees or lower."
End Sub

Private Sub cmdChart_Click()
picResultswax.Cls
I = 1
picResultswax.Print "Old, Slush, and Icy snow Wax Chart"
picResultswax.Print "Temp"; Tab(10); "Speed at all costs"; Tab(34); "Medium"; Tab(45); "Inexpensive"
Do While Not I = 24  'prints out complete wax chart
    picResultswax.Print Celsius(I); Tab(10); Fast(I); Tab(34); Medium(I); Tab(45); Slow(I)
    I = I + 1
Loop
picResultswax.Print ""  'prints a message to help user understand the waxing instructions
picResultswax.Print "***** If the temapture 6 appeared you should"
picResultswax.Print "use the wax for 6 degrees or or higher tempatures"
picResultswax.Print " and if -16 appeares use the wax for -16 degrees or lower."
End Sub

Private Sub cmdConvert_Click()
Dim F As Single, C As Single
    'converts fahrnheit into Celsius
    picTemp.Cls
    F = txtTemp
    C = (5 / 9) * (F - 32)
    picTemp.Print F; " degrees fahrnheit ="; FormatNumber(C, 2); "degrees centigrade"
End Sub

Private Sub cmdReturnmain_Click()
    OldSnowForm.Hide    'brings user to MainForm1
    MainForm1.Show
End Sub

Private Sub cmdReturnprevious2_Click()
    SnowForm.Show  'returns user to previous form
    OldSnowForm.Hide
End Sub

Private Sub cmdWax_Click()
picResultswax.Cls
Temp = InputBox("Enter the air tempature in celsius", "New And Fine Snow")
picResultswax.Print "Temp"; Tab(10); "Speed At All Cost"; Tab(34); "Medium"; Tab(45); "Economical"
Select Case Temp    'user enters a tempature and comapres it if statements untilit fits an equatuion than prints wax types for that tempature
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
    picResultswax.Print ""
    picResultswax.Print "Cerra F should only be applied  in situations"
    picResultswax.Print "when speed is is crucial, such as race day because"
    picResultswax.Print "of its extreme cost. Using the first wax listed under"
    picResultswax.Print "speed at all costs and not applying Cerra F should give"
    picResultswax.Print "most everyone the speed they are looking for."
End If
End Sub

Private Sub Form_Load()
Open NewSnowForm.Path & "Oldsnow.txt" For Input As #1 'reads in file as arraysOpen "M:\CS130\VB-Project\" For Input As #1
For I = 1 To 23
    Input #1, Celsius(I), Fast(I), Medium(I), Slow(I)
Next I
Close #1  'closes file
End Sub
