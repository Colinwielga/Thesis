VERSION 5.00
Begin VB.Form frmabortion 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Abortion"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Compute option"
      Height          =   2055
      Left            =   240
      TabIndex        =   15
      Top             =   8760
      Width           =   4215
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit now"
      Height          =   1455
      Left            =   10320
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   14
      Top             =   9000
      Width           =   2775
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Move on to the next page"
      Height          =   2055
      Left            =   5400
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   13
      Top             =   8760
      Width           =   4215
   End
   Begin VB.OptionButton absharpton 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Abortion is wrong, but let women choose."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7920
      Width           =   13335
   End
   Begin VB.OptionButton ablieberman 
      BackColor       =   &H00C0C0FF&
      Caption         =   "The decision to have an abortion should be left to a woman, her doctor, and her god."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   13335
   End
   Begin VB.OptionButton abkucinich 
      BackColor       =   &H00C0C0FF&
      Caption         =   $"frmabortion.frx":0000
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6720
      Width           =   13455
   End
   Begin VB.OptionButton abkerry 
      BackColor       =   &H00C0C0FF&
      Caption         =   "There should be no criminalization of a woman's right to choose."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   13335
   End
   Begin VB.OptionButton abgraham 
      BackColor       =   &H00C0C0FF&
      Caption         =   "The ban on military base abortions should not be maintained."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   13335
   End
   Begin VB.OptionButton abgephardt 
      BackColor       =   &H00C0C0FF&
      Caption         =   "There should be no partial birth abortions, but no barring of transporting minors to get  abortions."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   13215
   End
   Begin VB.OptionButton abedwards 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Partial birth abortions should not be banned."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   12975
   End
   Begin VB.OptionButton abbush 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Partial birth abortions should be banned, and the reduction of abortions overall should be encouraged via adoption and abstinence."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   13335
   End
   Begin VB.OptionButton abmoseley 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Women should all the say in if they have an abortion or not."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   13335
   End
   Begin VB.OptionButton abnader 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Government should have no role in the issue of abortion."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   13215
   End
   Begin VB.OptionButton abclark 
      BackColor       =   &H00C0C0FF&
      Caption         =   "My view is pro-choice; you should support the rights of women."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   13335
   End
   Begin VB.OptionButton abdean 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Women should have the right to choose."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   13215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Please click next to the statement that you most agree with concerning abortion."
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmabortion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Choosing A Candidate (M:\My Documents\ChoosingACandidate.vbp)
'frmabortion(frmabortion)
'Elaina Reinke
'October 30, 2003
'This program is for a user who wants to know which nominee for President,
'if the election was this year, they would choose.
'The program compares values of the user with values of a candidate
'and computes which candidate the user has the most similar values of
'frmabortion is a form that asks the user what they feel about abortion.
'the compute button assigns a number value to that option they choose from
'an array and adds it to the running sum which will be used at the end


Option Explicit

Private Sub abbush_Click()
'Enable "Compute Option" button after selection has been made
If abbush = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub abclark_Click()
'Enable "Compute Option" button after selection has been made
If abclark = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub


Private Sub abdean_Click()
'Enable "Compute Option" button after selection has been made
If abdean = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub abedwards_Click()
'Enable "Compute Option" button after selection has been made
If abedwards = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub abgephardt_Click()
'Enable "Compute Option" button after selection has been made
If abgephardt = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub abgraham_Click()
'Enable "Compute Option" button after selection has been made
If abgraham = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub abkerry_Click()
'Enable "Compute Option" button after selection has been made
If abkerry = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub abkucinich_Click()
'Enable "Compute Option" button after selection has been made
If abkucinich = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub ablieberman_Click()
'Enable "Compute Option" button after selection has been made
If ablieberman = True Then
    cmdcompute.Enabled = True
End If
End Sub

Private Sub abmoseley_Click()
'Enable "Compute Option" button after selection has been made
If abmoseley = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub abnader_Click()
'Enable "Compute Option" button after selection has been made
If abnader = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub absharpton_Click()
'Enable "Compute Option" button after selection has been made
If absharpton = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub cmdcompute_Click()
'Disable the "Move on to the next page" button
cmdnext.Enabled = False
'Determine which option the user has chosen
'and set it equal to the number in the array in which they have been ordered
'they have been ordered from most Democratic (1) through most Republican (12)
sum = 0

If abdean = True Then
    sum = sum + candvalues(1)
        ElseIf abclark = True Then
            sum = sum + candvalues(2)
        ElseIf abnader = True Then
            sum = sum + candvalues(3)
        ElseIf abmoseley = True Then
            sum = sum + candvalues(4)
        ElseIf abbush = True Then
            sum = sum + candvalues(5)
        ElseIf abedwards = True Then
            sum = sum + candvalues(6)
        ElseIf abgephardt = True Then
            sum = sum + candvalues(7)
        ElseIf abgraham = True Then
            sum = sum + candvalues(8)
        ElseIf abkerry = True Then
            sum = sum + candvalues(9)
        ElseIf abkucinich = True Then
            sum = sum + candvalues(10)
        ElseIf ablieberman = True Then
            sum = sum + candvalues(11)
        ElseIf absharpton = True Then
            sum = sum + candvalues(12)
        End If

'disable the "compute option" button and
'Enable the "Move on to the next page" button
cmdcompute.Enabled = False
cmdnext.Enabled = True
End Sub

Private Sub cmdnext_Click()
'hide this form and show the next one while enabling the "compute"
'button on the next page and disabling the "next" button
frmabortion.Hide
frmaffact.Show
cmdcompute.Enabled = True
cmdnext.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()
'load the file into the array
PATH = "M:\My Documents\Reinke\"
Open PATH & "candidatevalues.txt" For Input As #1
For i = 1 To 12
    Input #1, candvalues(i)
Next i
Close #1
'do these in form load so that they are already happening when you
'open the program, so you do not have to push any buttons to open
'the file to the array and input it
cmdcompute.Enabled = False
cmdnext.Enabled = False

End Sub
