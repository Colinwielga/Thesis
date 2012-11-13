VERSION 5.00
Begin VB.Form frmdrugs 
   BackColor       =   &H0080FF80&
   Caption         =   "Drugs"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit now"
      Height          =   975
      Left            =   9480
      TabIndex        =   15
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Move on to next page"
      Height          =   2055
      Left            =   4560
      TabIndex        =   14
      Top             =   7920
      Width           =   4095
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Compute Option"
      Height          =   2055
      Left            =   240
      TabIndex        =   13
      Top             =   7920
      Width           =   4095
   End
   Begin VB.OptionButton drsharpton 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The policy on drug offenses should stay the same."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   12375
   End
   Begin VB.OptionButton drlieberman 
      BackColor       =   &H00C0FFFF&
      Caption         =   "There should be increased penalites for drug offenses."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   12375
   End
   Begin VB.OptionButton drkucinich 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rehabilitation should be emphasized over incarceration."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   12375
   End
   Begin VB.OptionButton drkerry 
      BackColor       =   &H00C0FFFF&
      Caption         =   "No spending international development funds on drug control."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   12375
   End
   Begin VB.OptionButton drgraham 
      BackColor       =   &H00C0FFFF&
      Caption         =   "There should not be an increase in penalties for drug offenses.  International development funds should be spent on drug control."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   12375
   End
   Begin VB.OptionButton drgephardt 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"choosingacandidate.frx":0000
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   12375
   End
   Begin VB.OptionButton dredwards 
      BackColor       =   &H00C0FFFF&
      Caption         =   "I am against increased penalties for drug offenses."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   12375
   End
   Begin VB.OptionButton drbush 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"choosingacandidate.frx":009B
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   12375
   End
   Begin VB.OptionButton drmoseley 
      BackColor       =   &H00C0FFFF&
      Caption         =   "There should be no spending of international development funds on drug control."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   12375
   End
   Begin VB.OptionButton drnader 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"choosingacandidate.frx":012A
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   12375
   End
   Begin VB.OptionButton drclark 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Federal funds should not be increased against the war on drugs."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   12375
   End
   Begin VB.OptionButton drdean 
      BackColor       =   &H00C0FFFF&
      Caption         =   "I do not have much of a stance on the issue of drugs."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   12375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Please click next to the option you most agree with concerning drugs."
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmdrugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ChoosingACandidate(ChoosingACandidate)
'frm drugs(frmdrugs)
'Elaina Reinke
'October 30,2003
'this form asks the user what they feel about drugs and adds the assigned
'value to the running sum

Private Sub cmdcompute_Click()
'Disable the "Move on to the next page" button
cmdnext.Enabled = False
'Determine which option the user has chosen
'set the option equal to its value from 1 to 12 in the array

    
If drdean = True Then
    sum = sum + candvalues(1)
        ElseIf drclark = True Then
            sum = sum + candvalues(2)
        ElseIf drnader = True Then
            sum = sum + candvalues(3)
        ElseIf drmoseley = True Then
            sum = sum + candvalues(4)
        ElseIf drbush = True Then
            sum = sum + candvalues(5)
        ElseIf dredwards = True Then
           sum = sum + candvalues(6)
        ElseIf drgephardt = True Then
            sum = sum + candvalues(7)
        ElseIf drgraham = True Then
            sum = sum + candvalues(8)
        ElseIf drkerry = True Then
            sum = sum + candvalues(9)
        ElseIf drkucinich = True Then
            sum = sum + candvalues(10)
        ElseIf drlieberman = True Then
            sum = sum + candvalues(11)
        ElseIf drsharpton = True Then
            sum = sum + candvalues(12)
    End If

'disable the "compute option" button and
'Enable the "Move on to the next page" button
cmdcompute.Enabled = False
cmdnext.Enabled = True
End Sub

Private Sub cmdnext_Click()
'hide this form and show the next
frmdrugs.Hide
frmeducation.Show
cmdcompute.Enabled = True
cmdnext.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub drbush_Click()
'Enable "Compute Option" button after selection has been made
If drbush = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drclark_Click()
'Enable "Compute Option" button after selection has been made
If drclark = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drdean_Click()
'Enable "Compute Option" button after selection has been made
If drdean = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub dredwards_Click()
'Enable "Compute Option" button after selection has been made
If dredwards = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drgephardt_Click()
'Enable "Compute Option" button after selection has been made
If drgephardt = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drgraham_Click()
'Enable "Compute Option" button after selection has been made
If drgraham = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drkerry_Click()
'Enable "Compute Option" button after selection has been made
If drkerry = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drkucinich_Click()
'Enable "Compute Option" button after selection has been made
If drkucinich = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drlieberman_Click()
'Enable "Compute Option" button after selection has been made
If drlieberman = True Then
    cmdcompute.Enabled = True
End If
End Sub

Private Sub drmoseley_Click()
'Enable "Compute Option" button after selection has been made
If drmoseley = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drnader_Click()
'Enable "Compute Option" button after selection has been made
If drnader = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub drsharpton_Click()
'Enable "Compute Option" button after selection has been made
If drsharpton = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub Form_Load()
'putting these instructions in form load makes the program perform them
'without the need for a button specifically for that purpose
cmdcompute.Enabled = False
cmdnext.Enabled = False
Open PATH & "candidatevalues.txt" For Input As #1
For i = 1 To 12
    Input #1, candvalues(i)
Next i
Close #1
End Sub
