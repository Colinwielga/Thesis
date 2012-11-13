VERSION 5.00
Begin VB.Form frmaffact 
   BackColor       =   &H0080FFFF&
   Caption         =   "Affirmative Action"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit now"
      Height          =   1095
      Left            =   9840
      TabIndex        =   15
      Top             =   9000
      Width           =   2775
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Move on to the next page"
      Height          =   2055
      Left            =   4800
      TabIndex        =   14
      Top             =   8520
      Width           =   4215
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Compute Option"
      Height          =   2055
      Left            =   240
      TabIndex        =   13
      Top             =   8520
      Width           =   4215
   End
   Begin VB.OptionButton afsharpton 
      BackColor       =   &H0080FF80&
      Caption         =   "Affirmative action should continue, and the discussion of racial equality should be explored more."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   12735
   End
   Begin VB.OptionButton aflieberman 
      BackColor       =   &H0080FF80&
      Caption         =   "Affirmative action is necessary now, but I hope to phase it out by 2010."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   12735
   End
   Begin VB.OptionButton afkucinich 
      BackColor       =   &H0080FF80&
      Caption         =   "Affirmative action is necessary and right and must be preserved."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   12735
   End
   Begin VB.OptionButton afkerry 
      BackColor       =   &H0080FF80&
      Caption         =   "The act of hiring according to affirmative action with use of federal funds should not be banned."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   12735
   End
   Begin VB.OptionButton afgraham 
      BackColor       =   &H0080FF80&
      Caption         =   "I do not support affirmative action hiring with use of federal funds."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   12735
   End
   Begin VB.OptionButton afgephardt 
      BackColor       =   &H0080FF80&
      Caption         =   "There should be preferential treatment by race in college admissions."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   12615
   End
   Begin VB.OptionButton afedwards 
      BackColor       =   &H0080FF80&
      Caption         =   "Affirmative action was needed 40 years ago and is still needed today."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   12735
   End
   Begin VB.OptionButton afbush 
      BackColor       =   &H0080FF80&
      Caption         =   "I support affirmative action, but I don't support quotas or preferences.  We should reach out to minorities without quotas."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   12735
   End
   Begin VB.OptionButton afmoseley 
      BackColor       =   &H0080FF80&
      Caption         =   "I believe affirmative action should be supported with the use of federal funds."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   12735
   End
   Begin VB.OptionButton afnader 
      BackColor       =   &H0080FF80&
      Caption         =   "I support affirmative action, and wish to implement a Truth and Reconciliation Commision for Native Americans."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   12615
   End
   Begin VB.OptionButton afclark 
      BackColor       =   &H0080FF80&
      Caption         =   "I am a proponent and strong supporter of affirmative action, diversity, and multiculturalism."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   12615
   End
   Begin VB.OptionButton afdean 
      BackColor       =   &H0080FF80&
      Caption         =   "I support affirmative action and want to work toward an end to racial profiling."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Please click next to the option that you most agree with concerning affirmative action."
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmaffact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ChoosingACandidate(ChoosingACandidate)
'frm affirmative action(frmaffact)
'Elaina Reinke
'October 30,2003
'this form asks the user which option they agree with most concerning
'affirmative action.  The compute button assigns a value from the array
'to the option the user chooses and adds it to the running sum

Private Sub afbush_Click()
'Enable "Compute Option" button after selection has been made
If afbush = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afclark_Click()
'Enable "Compute Option" button after selection has been made
If afclark = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afdean_Click()
'Enable "Compute Option" button after selection has been made
If afdean = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afedwards_Click()
'Enable "Compute Option" button after selection has been made
If afedwards = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afgephardt_Click()
'Enable "Compute Option" button after selection has been made
If afgephardt = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afgraham_Click()
'Enable "Compute Option" button after selection has been made
If afgraham = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afkerry_Click()
'Enable "Compute Option" button after selection has been made
If afkerry = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afkucinich_Click()
'Enable "Compute Option" button after selection has been made
If afkucinich = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub aflieberman_Click()
'Enable "Compute Option" button after selection has been made
If aflieberman = True Then
    cmdcompute.Enabled = True
End If
End Sub

Private Sub afmoseley_Click()
'Enable "Compute Option" button after selection has been made
If afmoseley = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afnader_Click()
'Enable "Compute Option" button after selection has been made
If afnader = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub afsharpton_Click()
'Enable "Compute Option" button after selection has been made
If afsharpton = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub cmdcompute_Click()

'Dim i As Integer, sum As Single
'Disable the "Move on to the next page" button
cmdnext.Enabled = False
'Dim candvalues(1 To 12) As Integer
'Declare variables

'Determine which option the user has chosen
'and set it equal to its value from 1 to 12, 1 being most Democratic
'and 12 being most Republican

If afdean = True Then
    sum = sum + candvalues(1)
    ElseIf afclark = True Then
            sum = sum + candvalues(2)
    ElseIf afnader = True Then
            sum = sum + candvalues(3)
    ElseIf afmoseley = True Then
            sum = sum + candvalues(4)
    ElseIf afbush = True Then
            sum = sum + candvalues(5)
    ElseIf afedwards = True Then
            sum = sum + candvalues(6)
    ElseIf afgephardt = True Then
            sum = sum + candvalues(7)
    ElseIf afgraham = True Then
            sum = sum + candvalues(8)
    ElseIf afkerry = True Then
            sum = sum + candvalues(9)
    ElseIf afkucinich = True Then
            sum = sum + candvalues(10)
    ElseIf aflieberman = True Then
            sum = sum + candvalues(11)
    ElseIf afsharpton = True Then
            sum = sum + candvalues(12)
End If
'disable the "compute option" button and
'Enable the "Move on to the next page" button
cmdcompute.Enabled = False
cmdnext.Enabled = True
End Sub

Private Sub cmdnext_Click()
'hide this form and show the next one
frmaffact.Hide
frmgayrights.Show
cmdcompute.Enabled = True
cmdnext.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub


Private Sub Form_Load()
Open PATH & "candidatevalues.txt" For Input As #1
'Dim candvalues(1 To 12) As Integer
'Dim i As Integer

For i = 1 To 12
    Input #1, candvalues(i)
Next i
Close #1

'put these instructions in form load because then the program
'will do it without needing to push any buttons
cmdcompute.Enabled = False
cmdnext.Enabled = False

End Sub
