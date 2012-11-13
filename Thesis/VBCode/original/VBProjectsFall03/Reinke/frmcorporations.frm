VERSION 5.00
Begin VB.Form frmcorporations 
   BackColor       =   &H0080C0FF&
   Caption         =   "Corporations"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit now"
      Height          =   975
      Left            =   9720
      TabIndex        =   15
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Move to the next page"
      Height          =   2055
      Left            =   4800
      TabIndex        =   14
      Top             =   7920
      Width           =   4215
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Compute Option"
      Height          =   2055
      Left            =   240
      TabIndex        =   13
      Top             =   7920
      Width           =   4215
   End
   Begin VB.OptionButton cosharpton 
      BackColor       =   &H008080FF&
      Caption         =   "People are told to serve their country.  Corporations should do the same."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7200
      Width           =   12615
   End
   Begin VB.OptionButton colieberman 
      BackColor       =   &H008080FF&
      Caption         =   "Foxes guard the foxes and middle-class hens get plucked."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6600
      Width           =   12615
   End
   Begin VB.OptionButton cokucinich 
      BackColor       =   &H008080FF&
      Caption         =   "Democracy fails without corporate regulations."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   12615
   End
   Begin VB.OptionButton cokerry 
      BackColor       =   &H008080FF&
      Caption         =   "Efforts should be made to democratize the process of corporate boards."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   12615
   End
   Begin VB.OptionButton cograham 
      BackColor       =   &H008080FF&
      Caption         =   "The rules on personal bankruptcy should be restricted."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   12615
   End
   Begin VB.OptionButton cogephardt 
      BackColor       =   &H008080FF&
      Caption         =   "Greed can kill democracy and capitalism."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   12615
   End
   Begin VB.OptionButton coedwards 
      BackColor       =   &H008080FF&
      Caption         =   "Tax incentives should be awarded to companies to keep jobs in America."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   12615
   End
   Begin VB.OptionButton cobush 
      BackColor       =   &H008080FF&
      Caption         =   $"frmcorporations.frx":0000
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   12495
   End
   Begin VB.OptionButton comoseley 
      BackColor       =   &H008080FF&
      Caption         =   "There should be an increase in tobacco restrictions."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   12615
   End
   Begin VB.OptionButton conader 
      BackColor       =   &H008080FF&
      Caption         =   $"frmcorporations.frx":008E
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   12615
   End
   Begin VB.OptionButton coclark 
      BackColor       =   &H008080FF&
      Caption         =   "There should be more government involvement in businesses."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   12615
   End
   Begin VB.OptionButton codean 
      BackColor       =   &H008080FF&
      Caption         =   "The amount of federal support payments huge megafarms receive should be restricted."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "Please click next to the option that you most agree with concerning corporations."
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmcorporations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ChoosingACandidate(ChoosingACandidate)
'frm corporations(frmcorporations)
'Elaina Reinke
'October 30, 2003
'this form asks the user which option they agree with concerning
'corporations and adds its value from the array to the running sum


Private Sub cmdcompute_Click()
'Disable the "Move on to the next page" button
cmdnext.Enabled = False
'Determine which option the user has chosen
'and set it equal to its value


If codean = True Then
    sum = sum + candvalues(1)
        ElseIf coclark = True Then
            sum = sum + candvalues(2)
        ElseIf conader = True Then
            sum = sum + candvalues(3)
        ElseIf comoseley = True Then
            sum = sum + candvalues(4)
        ElseIf cobush = True Then
            sum = sum + candvalues(5)
        ElseIf coedwards = True Then
            sum = sum + candvalues(6)
        ElseIf cogephardt = True Then
            sum = sum + candvalues(7)
        ElseIf cograham = True Then
            sum = sum + candvalues(8)
        ElseIf cokerry = True Then
            sum = sum + candvalues(9)
        ElseIf cokucinich = True Then
            sum = sum + candvalues(10)
        ElseIf colieberman = True Then
            sum = sum + candvalues(11)
        ElseIf cosharpton = True Then
            sum = sum + candvalues(12)
    End If



'disable the "compute option" button and
'Enable the "Move on to the next page" button
cmdcompute.Enabled = False
cmdnext.Enabled = True
End Sub

Private Sub cmdnext_Click()
'hide this form and show the next one
frmcorporations.Hide
frmenvironment.Show
cmdcompute.Enabled = True
cmdnext.Enabled = False
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cobush_Click()
'Enable "Compute Option" button after selection has been made
If cobush = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub coclark_Click()
'Enable "Compute Option" button after selection has been made
If coclark = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub codean_Click()
'Enable "Compute Option" button after selection has been made
If codean = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub coedwards_Click()
'Enable "Compute Option" button after selection has been made
If coedwards = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub cogephardt_Click()
'Enable "Compute Option" button after selection has been made
If cogephardt = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub cograham_Click()
'Enable "Compute Option" button after selection has been made
If cograham = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub cokerry_Click()
'Enable "Compute Option" button after selection has been made
If cokerry = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub cokucinich_Click()
'Enable "Compute Option" button after selection has been made
If cokucinich = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub colieberman_Click()
'Enable "Compute Option" button after selection has been made
If colieberman = True Then
    cmdcompute.Enabled = True
End If
End Sub

Private Sub comoseley_Click()
'Enable "Compute Option" button after selection has been made
If comoseley = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub conader_Click()
'Enable "Compute Option" button after selection has been made
If conader = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub cosharpton_Click()
'Enable "Compute Option" button after selection has been made
If cosharpton = True Then
    cmdcompute.Enabled = True
Else: cmdcompute.Enabled = False
End If
End Sub

Private Sub Form_Load()
'putting this in form load makes the program do it without
'the need to push a button
cmdcompute.Enabled = False
cmdnext.Enabled = False

Open PATH & "candidatevalues.txt" For Input As #1
For i = 1 To 12
    Input #1, candvalues(i)
Next i
Close #1
End Sub
