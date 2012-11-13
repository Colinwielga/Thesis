VERSION 5.00
Begin VB.Form frmDoctor 
   BackColor       =   &H00400000&
   Caption         =   "Doctor"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14100
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
   ScaleHeight     =   10950
   ScaleWidth      =   14100
   Begin VB.CommandButton cmdExploreCosts 
      BackColor       =   &H0080FFFF&
      Caption         =   "Explore Costs"
      Height          =   800
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9720
      Width           =   2520
   End
   Begin VB.CommandButton cmdDisposableIncome 
      BackColor       =   &H0080FFFF&
      Caption         =   "Disposable Income"
      Height          =   800
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9720
      Width           =   2520
   End
   Begin VB.CommandButton cmdTaxBracket 
      BackColor       =   &H0080FFFF&
      Caption         =   "Find Tax Bracket"
      Height          =   800
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   2520
   End
   Begin VB.CommandButton cmdtax 
      BackColor       =   &H0080FFFF&
      Caption         =   "Find Tax on income "
      Height          =   800
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7800
      Width           =   2520
   End
   Begin VB.CommandButton cmdTaxPercent 
      BackColor       =   &H0080FFFF&
      Caption         =   "Tax as a % of income"
      Height          =   800
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8760
      Width           =   2520
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "Switch Professions"
      Height          =   800
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   2520
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H0080FFFF&
      Caption         =   "View Salary"
      Height          =   800
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   2520
   End
   Begin VB.TextBox txtLastName 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   2520
   End
   Begin VB.TextBox txtSex 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   615
   End
   Begin VB.PictureBox picDoctor 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   3240
      ScaleHeight     =   5355
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   1920
      Width           =   8415
   End
   Begin VB.PictureBox picWelcome 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   2520
      ScaleHeight     =   795
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   600
      Width           =   9975
   End
   Begin VB.Label lblMyBudget 
      BackColor       =   &H00400000&
      Caption         =   "My Budget Software  "
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblSex 
      BackColor       =   &H00400000&
      Caption         =   "Please enter sex:  M or F"
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblLastName 
      BackColor       =   &H00400000&
      Caption         =   "Please Enter Last Name:"
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "frmDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is unique to the profession of a Doctor. The user obtains the average salary of a doctor, and is
'able to determine vital information regarding taxes. This form also leads the user to an expense database.

'Ensures that all variables are declared and serves as a spell checker for variables
Option Explicit
Dim tax As Single

'Reveal salary and computational buttons
Private Sub cmdContinue_Click()

    picDoctor.Visible = True
    cmdtax.Visible = True
    cmdTaxBracket.Visible = True
    cmdTaxPercent.Visible = True
    cmdDisposableIncome.Visible = True
    cmdExploreCosts.Visible = True

    Dim Pos As Integer
    Dim Found As Boolean
    Dim LastName As String
    Dim Sex As String
    Dim S As String
    
    LastName = txtLastName.Text
    Sex = txtSex.Text
    
    S = "Dr."
       
    picWelcome.Cls
    picDoctor.Cls
    picWelcome.Print "Welcome " & S & " " & LastName & ""
    picWelcome.Print "Your starting salary as a Doctor is " & FormatCurrency(yourSalary) & ""
    picDoctor.Picture = LoadPicture(App.Path & "\images\" & "Doctor.jpg")
    
    lblLastName.Visible = False
    txtLastName.Visible = False
    lblSex.Visible = False
    txtSex.Visible = False
End Sub

'Return to previous form
Private Sub cmdBack_Click()
    frmDoctor.Visible = False
    frmStart.Visible = True
    picDoctor.Visible = False
    lblLastName.Visible = True
    txtLastName.Visible = True
    lblSex.Visible = True
    txtSex.Visible = True
    cmdtax.Visible = False
    cmdTaxBracket.Visible = False
    cmdTaxPercent.Visible = False
    cmdDisposableIncome.Visible = False
    cmdExploreCosts.Visible = False
End Sub

'Load form
Private Sub Form_Load()
    picDoctor.Visible = False
    cmdtax.Visible = False
    cmdTaxBracket.Visible = False
    cmdTaxPercent.Visible = False
    cmdDisposableIncome.Visible = False
    cmdExploreCosts.Visible = False
End Sub

'Calculate tax based on salary
Private Sub cmdtax_Click()

    picWelcome.Cls
    
    If yourSalary >= 0 And yourSalary <= 8375 Then
            tax = (yourSalary * 0.1)
            picWelcome.Print "You have to pay the Federal Government " & FormatCurrency(tax) & " in taxes!"
        ElseIf yourSalary > 8375 And yourSalary <= 34000 Then
            tax = ((8375 - 0) * 0.1) + ((yourSalary - 8375) * 0.15)
            picWelcome.Print "You have to pay the Federal Government " & FormatCurrency(tax) & " in taxes!"
        ElseIf yourSalary > 34000 And yourSalary <= 82400 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((yourSalary - 34000) * 0.25)
            picWelcome.Print "You have to pay the Federal Government " & FormatCurrency(tax) & " in taxes!"
        ElseIf yourSalary > 82400 And yourSalary <= 171850 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((yourSalary - 82400) * 0.28)
            picWelcome.Print "You have to pay the Federal Government " & FormatCurrency(tax) & " in taxes!"
        ElseIf yourSalary > 171850 And yourSalary <= 373650 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((171850 - 82400) * 0.28) + ((yourSalary - 171850) * 0.33)
            picWelcome.Print "You have to pay the Federal Government " & FormatCurrency(tax) & " in taxes!"
        ElseIf yourSalary > 373650 Then
           tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((171850 - 82400) * 0.28) + ((373650 - 171850) * 0.33) + ((yourSalary - 373650) * 0.35)
            picWelcome.Print "You have to pay the Federal Government " & FormatCurrency(tax) & " in taxes!"
        Else
            picWelcome.Print "Please declare salary"
    End If
End Sub

'Find tax bracket using calculation
Private Sub cmdTaxBracket_Click()
    
    picWelcome.Cls
    
    If yourSalary >= 0 And yourSalary <= 8375 Then
            picWelcome.Print "Your tax bracket is " & FormatPercent(0.1, 0) & ""
        ElseIf yourSalary > 8375 And yourSalary <= 34000 Then
            picWelcome.Print "Your tax bracket is " & FormatPercent(0.15) & ""
        ElseIf yourSalary > 34000 And yourSalary <= 82400 Then
            picWelcome.Print "Your tax bracket is " & FormatPercent(0.25, 0) & ""
        ElseIf yourSalary > 82400 And yourSalary <= 171850 Then
            picWelcome.Print "Your tax bracket is " & FormatPercent(0.28, 0) & ""
        ElseIf yourSalary > 171850 And yourSalary <= 373650 Then
            picWelcome.Print "Your tax bracket is " & FormatPercent(0.33, 0) & ""
        ElseIf yourSalary > 373650 Then
            picWelcome.Print "Your tax bracket is " & FormatPercent(0.35, 0) & ""
        Else
            picWelcome.Print "Please declare salary"
    End If
            

End Sub

'Calculate Percentage of taxes
Private Sub cmdTaxPercent_Click()
    Dim taxPercent As Single
    
     picWelcome.Cls
        
    If yourSalary >= 0 And yourSalary <= 8375 Then
            tax = (yourSalary * 0.1)
            taxPercent = tax / yourSalary
            picWelcome.Print "Your tax is " & FormatPercent(taxPercent, 2) & " of your salary based income!"
        ElseIf yourSalary > 8375 And yourSalary <= 34000 Then
            tax = ((8375 - 0) * 0.1) + ((yourSalary - 8375) * 0.15)
            taxPercent = tax / yourSalary
            picWelcome.Print "Your tax is " & FormatPercent(taxPercent, 2) & " of your salary based income!"
        ElseIf yourSalary > 34000 And yourSalary <= 82400 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((yourSalary - 34000) * 0.25)
            taxPercent = tax / yourSalary
            picWelcome.Print "Your tax is " & FormatPercent(taxPercent, 2) & " of your salary based income!"
        ElseIf yourSalary > 82400 And yourSalary <= 171850 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((yourSalary - 82400) * 0.28)
            taxPercent = tax / yourSalary
            picWelcome.Print "Your tax is " & FormatPercent(taxPercent, 2) & " of your salary based income!"
        ElseIf yourSalary > 171850 And yourSalary <= 373650 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((171850 - 82400) * 0.28) + ((yourSalary - 171850) * 0.33)
            taxPercent = tax / yourSalary
            picWelcome.Print "Your tax is " & FormatPercent(taxPercent, 2) & " of your salary based income!"
        ElseIf yourSalary > 373650 Then
           tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((171850 - 82400) * 0.28) + ((373650 - 171850) * 0.33) + ((yourSalary - 373650) * 0.35)
            taxPercent = tax / yourSalary
            picWelcome.Print "Your tax is " & FormatPercent(taxPercent, 2) & " of your salary based income!"
        Else
            picWelcome.Print "Please declare salary"
    End If
End Sub

'Calculate disposable income
Private Sub cmdDisposableIncome_Click()
    Dim DisposableIncome As Single
    
    picWelcome.Cls
    
    If yourSalary >= 0 And yourSalary <= 8375 Then
            tax = (yourSalary * 0.1)
           
        ElseIf yourSalary > 8375 And yourSalary <= 34000 Then
            tax = ((8375 - 0) * 0.1) + ((yourSalary - 8375) * 0.15)
            
        ElseIf yourSalary > 34000 And yourSalary <= 82400 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((yourSalary - 34000) * 0.25)
            
        ElseIf yourSalary > 82400 And yourSalary <= 171850 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((yourSalary - 82400) * 0.28)
           
        ElseIf yourSalary > 171850 And yourSalary <= 373650 Then
            tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((171850 - 82400) * 0.28) + ((yourSalary - 171850) * 0.33)
            
        ElseIf yourSalary > 373650 Then
           tax = ((8375 - 0) * 0.1) + ((34000 - 8375) * 0.15) + ((82400 - 34000) * 0.25) + ((171850 - 82400) * 0.28) + ((373650 - 171850) * 0.33) + ((yourSalary - 373650) * 0.35)
            
    End If
    
    DisposableIncome = yourSalary - tax
    picWelcome.Print "Your disposable income is " & FormatCurrency(DisposableIncome) & ""
            
        
End Sub

'Load file into arrays and lead user to cost form
Private Sub cmdExploreCosts_Click()
    frmDoctor.Visible = False
    frmCost.Visible = True
    
    Open App.Path & "\CostofLiving.txt" For Input As #1
    ctrRead = 0
    Do Until EOF(1)
        ctrRead = ctrRead + 1
        Input #1, UsState(ctrRead), OneAdult(ctrRead), OneAdultOneChild(ctrRead), TwoAdults(ctrRead), TwoAdultsOneChild(ctrRead), TwoAdultsTwoChildren(ctrRead)
    Loop
    
    Close #1
    
End Sub



















