VERSION 5.00
Begin VB.Form frmcrime 
   BackColor       =   &H000000FF&
   Caption         =   "Crime"
   ClientHeight    =   10665
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit now"
      Height          =   1215
      Left            =   9840
      TabIndex        =   15
      Top             =   8520
      Width           =   2775
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Move on to the next page"
      Height          =   2055
      Left            =   4680
      TabIndex        =   14
      Top             =   8280
      Width           =   4215
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Compute Option"
      Height          =   2055
      Left            =   240
      TabIndex        =   13
      Top             =   8280
      Width           =   4215
   End
   Begin VB.OptionButton crsharpton 
      BackColor       =   &H00FF0000&
      Caption         =   "I am opposed to the death penalty and believe it is no coincidence that the wealthy do not get executed."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   13215
   End
   Begin VB.OptionButton crlieberman 
      BackColor       =   &H00FF0000&
      Caption         =   "I favor the death penalty, even for minors.  More jails should be built to keep violent offenders for their full sentences."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   13215
   End
   Begin VB.OptionButton crkucinich 
      BackColor       =   &H00FF0000&
      Caption         =   "The federal death penalty should absolutely be terminated."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   13215
   End
   Begin VB.OptionButton crkerry 
      BackColor       =   &H00FF0000&
      Caption         =   "There should be no limit on death penalty appeals and no rejection of racial statistics in death penalty appeals."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   13215
   End
   Begin VB.OptionButton crgraham 
      BackColor       =   &H00FF0000&
      Caption         =   "There should be a limit on death penalty appeals, and racial statistics in death penalty appeals should be rejected."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   13215
   End
   Begin VB.OptionButton crgephardt 
      BackColor       =   &H00FF0000&
      Caption         =   "The death penalty should not be replaced with life imprisonment.  DNA testing should be required for all federal executions."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   13215
   End
   Begin VB.OptionButton credwards 
      BackColor       =   &H00FF0000&
      Caption         =   "There should be more funding and stricter sentencing for hate crimes."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   13215
   End
   Begin VB.OptionButton crbush 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmcrime.frx":0000
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   13215
   End
   Begin VB.OptionButton crmoseley 
      BackColor       =   &H00FF0000&
      Caption         =   "A limit should be put on product liability punitive damage awards."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   13215
   End
   Begin VB.OptionButton crnader 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmcrime.frx":00A7
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   13215
   End
   Begin VB.OptionButton crclark 
      BackColor       =   &H00FF0000&
      Caption         =   "Assault weapons should be banned for the general public."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   13215
   End
   Begin VB.OptionButton crdean 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmcrime.frx":0175
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   13095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Please click next to the option you most agree with concerning crime."
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmcrime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ChoosingACandidate(ChoosingACandidate)
'frm crime(frmcrime)
'Elaina Reinke
'October 30,2003
'this form asks the user what they feel about crime and adds the assigne
'value to the running sum

Private Sub cmdcompute_Click()
'Disable the "Move on to the next page" button
cmdnext.Enabled = False
'Determine which option the user has chosen
'and set it equal to its value

If crdean = True Then
   sum = sum + candvalues(1)
        ElseIf crclark = True Then
            sum = sum + candvalues(2)
        ElseIf crnader = True Then
            sum = sum + candvalues(3)
        ElseIf crmoseley = True Then
            sum = sum + candvalues(4)
        ElseIf crbush = True Then
            sum = sum + candvalues(5)
        ElseIf credwards = True Then
            sum = sum + candvalues(6)
        ElseIf crgephardt = True Then
            sum = sum + candvalues(7)
        ElseIf crgraham = True Then
            sum = sum + candvalues(8)
        ElseIf crkerry = True Then
            sum = sum + candvalues(9)
        ElseIf crkucinich = True Then
            sum = sum + candvalues(10)
        ElseIf crlieberman = True Then
            sum = sum + candvalues(11)
        ElseIf crsharpton = True Then
            sum = sum + candvalues(12)
    End If


'disable the "compute option" button and
'Enable the "Move on to the next page" button
cmdcompute.Enabled = False
cmdnext.Enabled = True
End Sub

Private Sub cmdnext_Click()
'hide this form and show the next
frmcrime.Hide
frmdrugs.Show
cmdcompute.Enabled = True
cmdnext.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub crbush_Click()
'Enable "Compute Option" button after selection has been made
If crbush = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crclark_Click()
'Enable "Compute Option" button after selection has been made
If crclark = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crdean_Click()
'Enable "Compute Option" button after selection has been made
If crdean = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub credwards_Click()
'Enable "Compute Option" button after selection has been made
If credwards = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crgephardt_Click()
'Enable "Compute Option button after selection has been made
If crgephardt = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crgraham_Click()
'Enable "Compute Option" button after selection has been made
If crgraham = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crkerry_Click()
'Enable "Compute Option" button after selection has been made
If crkerry = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crkucinich_Click()
'Enable "Compute Option" button after selection has been made
If crkucinich = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crlieberman_Click()
'Enable "Compute Option" button after selection has been made
If crlieberman = True Then
    cmdcompute.Enabled = True
End If
End Sub

Private Sub crmoseley_Click()
'Enable "Compute Option" button after selection has been made
If crmoseley = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crnader_Click()
'Enable "Compute Option" button after selection has been made
If crnader = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub crsharpton_Click()
'Enable "Compute Option" button after selection has been made
If crsharpton = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub Form_Load()
'form load enables these instructions without the need to
'push a button
'this is opening the file for the array we need in order
'to compute the data into the option selected
cmdcompute.Enabled = False
cmdnext.Enabled = False
Open PATH & "candidatevalues.txt" For Input As #1
For i = 1 To 12
    Input #1, candvalues(i)
Next i
Close #1
End Sub
