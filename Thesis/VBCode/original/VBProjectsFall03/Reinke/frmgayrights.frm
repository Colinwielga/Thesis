VERSION 5.00
Begin VB.Form frmgayrights 
   BackColor       =   &H00FF80FF&
   Caption         =   "Gay Rights"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit now"
      Height          =   1335
      Left            =   11160
      TabIndex        =   15
      Top             =   8760
      Width           =   2655
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Move on to the next  page"
      Height          =   2055
      Left            =   6000
      TabIndex        =   14
      Top             =   8520
      Width           =   4095
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Compute Option"
      Height          =   2055
      Left            =   840
      TabIndex        =   13
      Top             =   8520
      Width           =   4095
   End
   Begin VB.OptionButton gasharpton 
      BackColor       =   &H00FFFFC0&
      Caption         =   "I support gay rights.  Let people choose to sin or not.  They should be allowed to adopt."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7560
      Width           =   13815
   End
   Begin VB.OptionButton galieberman 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmgayrights.frx":0000
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   13815
   End
   Begin VB.OptionButton gakucinich 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Same-sex couples deserve equal domestic benefits, and gay adoptions should not be banned."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   13815
   End
   Begin VB.OptionButton gakerry 
      BackColor       =   &H00FFFFC0&
      Caption         =   "There should be no prohibition of same-sex marriages."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   13815
   End
   Begin VB.OptionButton gagraham 
      BackColor       =   &H00FFFFC0&
      Caption         =   "I think same-sex marriages should be prohibited."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   13815
   End
   Begin VB.OptionButton gagephardt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Gay adoptions in D.C. is okay."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   13815
   End
   Begin VB.OptionButton gaedwards 
      BackColor       =   &H00FFFFC0&
      Caption         =   "The government does not belong in anyone's bedroom, including gay bedrooms."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   13815
   End
   Begin VB.OptionButton gabush 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmgayrights.frx":0091
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   13815
   End
   Begin VB.OptionButton gamoseley 
      BackColor       =   &H00FFFFC0&
      Caption         =   "GLBT have the right to consensual gay sex.  I support same-sex marriages."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   13815
   End
   Begin VB.OptionButton ganader 
      BackColor       =   &H00FFFFC0&
      Caption         =   "I support gay marriage, and equal gay rights which includes  civil unions."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   13815
   End
   Begin VB.OptionButton gaclark 
      BackColor       =   &H00FFFFC0&
      Caption         =   "We should welcome people who want to serve in the military."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   13695
   End
   Begin VB.OptionButton gadean 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Equal rights should be expanded to same-sex couples and worplace discrimination should be banned."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   13815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Please click next to the option that you most agree with concerning gay rights."
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmgayrights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ChoosingACandidate(ChoosingACandidate)
'frm gay rights(frmgayrights)
'Elaina Reinke
'October 30,2003
'this form asks the user which option they agree with most concerning
'gay rights and adds the assigned value to the running sum

Private Sub cmdcompute_Click()
'Disable the "Move on to the next page" button
cmdnext.Enabled = False
'Determine which option the user has chosen
'set the option chosen equal to its section of the array


If gadean = True Then
    sum = sum + candvalues(1)
        ElseIf gaclark = True Then
            sum = sum + candvalues(2)
        ElseIf ganader = True Then
            sum = sum + candvalues(3)
        ElseIf gamoseley = True Then
            sum = sum + candvalues(4)
        ElseIf gabush = True Then
            sum = sum + candvalues(5)
        ElseIf gaedwards = True Then
            sum = sum + candvalues(6)
        ElseIf gagephardt = True Then
            sum = sum + candvalues(7)
        ElseIf gagraham = True Then
            sum = sum + candvalues(8)
        ElseIf gakerry = True Then
            sum = sum + candvalues(9)
        ElseIf gakucinich = True Then
            sum = sum + candvalues(10)
        ElseIf galieberman = True Then
            sum = sum + candvalues(11)
        ElseIf gasharpton = True Then
            sum = sum + candvalues(12)
    End If


'disable the "compute option" button and
'Enable the "Move on to the next page" button
cmdcompute.Enabled = False
cmdnext.Enabled = True
End Sub

Private Sub cmdnext_Click()
'hide this form and show the next
frmgayrights.Hide
frmcrime.Show
cmdcompute.Enabled = True
cmdnext.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()
'open the file data for the array in form load so we
'do not have to add a button to do so
cmdcompute.Enabled = False
cmdnext.Enabled = False
Open PATH & "candidatevalues.txt" For Input As #1
For i = 1 To 12
    Input #1, candvalues(i)
Next i
Close #1
End Sub

Private Sub gabush_Click()
'Enable "Compute Option" button after selection has been made
If gabush = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub gaclark_Click()
'Enable "Compute Option" button after selection has been made
If gaclark = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub gadean_Click()
'Enable "Compute Option" button after selection has been made
If gadean = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub gaedwards_Click()
'Enable "Compute Option" button after selection has been made
If gaedwards = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub gagephardt_Click()
'Enable "Compute Option" button after selection has been made
If gagephardt = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub gagraham_Click()
'Enable "Compute Option" button after selection has been made
If gagraham = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub gakerry_Click()
'Enable "Compute Option" button after selection has been made
If gakerry = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub gakucinich_Click()
'Enable "Compute Option" button after selection has been made
If gakucinich = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub galieberman_Click()
'Enable "Compute Option" button after selection has been made
If galieberman = True Then
    cmdcompute.Enabled = True
End If
End Sub

Private Sub gamoseley_Click()
'Enable "Compute Option" button after selection has been made
If gamoseley = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub ganader_Click()
'Enable "Compute Option" button after selection has been made
If ganader = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub gasharpton_Click()
'Enable "Compute Option" button after selection has been made
If gasharpton = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub
