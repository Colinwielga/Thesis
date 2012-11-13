VERSION 5.00
Begin VB.Form frmRules 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rules"
   ClientHeight    =   7440
   ClientLeft      =   6690
   ClientTop       =   4515
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   12495
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   480
      ScaleHeight     =   4455
      ScaleWidth      =   11055
      TabIndex        =   5
      Top             =   2280
      Width           =   11055
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return To Main Page"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FF80FF&
      Caption         =   "Search For Rules By Rule Number"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdDisplayReverse 
      BackColor       =   &H0080FF80&
      Caption         =   "Display Rules in Reverse Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FFFF80&
      Caption         =   "Read Rules"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RuleNumber(1 To 100) As Integer, Rule(1 To 100) As String
Dim CTR As Integer
'Family Feud
'frmRules
'Colin Hall and Andre Blaine
'March 15
    'This form will read the file into two parallel arrays,
    'search the file and print the rule based on the rule number given by the user,
    'sort the file and print the file in descending order,
    'return to the Main Page Form,
    'and will quit.

Private Sub cmdRead_Click()
    
    'This will clear ther information from picResults.
    picResults.Cls
    
    'This will open the file.
    Open App.Path & "\FamilyFeudRules.txt" For Input As #1
    
    'This will inform the user that the array reading has been completed.
    picResults.Print "The parallel array has been read successfully."
    picResults.Print " "
    picResults.Print "Rule Number"; Tab(32); "Rule"
    picResults.Print "*******************************************************************************************"
    
    'This will read the file into two parallel arrays and will display the array as found in the notepad.
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, RuleNumber(CTR), Rule(CTR)
        picResults.Print RuleNumber(CTR); Tab(16); Rule(CTR)
    Loop
    
    'This closes the file we read into an array.
    Close #1
    
    'This will make the Display Reverse and Search buttons able to be clicked.
    cmdDisplayReverse.Enabled = True
    cmdSearch.Enabled = True

End Sub
    
Private Sub cmdDisplayReverse_Click()

    Dim Pass As Integer, POS As Integer, TempRuleNumber As Integer, TempRule As String, I As Integer
    
    'This will clear ther information from picResults.
    picResults.Cls
    
    'This will sort the rule numbers and rules into reverse order.
    For Pass = 1 To CTR - 1
        For POS = 1 To CTR - Pass
            If RuleNumber(POS) < RuleNumber(POS + 1) Then
                TempRuleNumber = RuleNumber(POS)
                RuleNumber(POS) = RuleNumber(POS + 1)
                RuleNumber(POS + 1) = TempRuleNumber
                TempRule = Rule(POS)
                Rule(POS) = Rule(POS + 1)
                Rule(POS + 1) = TempRule
            End If
        Next POS
    Next Pass
    
    'This will print the rule number and rule headings.
    picResults.Print "Rule Number"; Tab(32); "Rule"
    picResults.Print "*******************************************************************************************"
    
    'This will print the rule numbers and rules in descending order.
    For I = 1 To CTR
        picResults.Print RuleNumber(I), Rule(I)
    Next I

End Sub

Private Sub cmdSearch_Click()
    
    Dim Found As Boolean, J As Integer, UserRule As Integer
    
    'This will clear the information from picResults.
    picResults.Cls
    
    'This will ask the user for a rule number ans assign it to UserRule.
    UserRule = InputBox("Enter the rule number you wish to see.", "Enter Number")
    
    'This will assign Found as false.
    Found = False
    
    'This will look for UserRule in the array and if found will print the rule number and rule.
    For J = 1 To CTR
        If RuleNumber(J) = UserRule Then
            picResults.Print "Rule Number"; Tab(32); "Rule"
            picResults.Print "*******************************************************************************************"
            picResults.Print RuleNumber(J), Rule(J)
            Found = True
        End If
    Next J
    
    'This will print a statement indicating if the UserRule does not exist in the array.
    If Not Found Then
        picResults.Print "Rule number "; UserRule; " does not exist."
    End If
    
End Sub

Private Sub cmdReturn_Click()

    'This button will hide the Creators form and will open the Main Page form.
    frmMainPage.Show
    frmCreators.Hide
    
End Sub

Private Sub cmdQuit_Click()

    'This button will end the Visual Basic Program.
    End

End Sub
