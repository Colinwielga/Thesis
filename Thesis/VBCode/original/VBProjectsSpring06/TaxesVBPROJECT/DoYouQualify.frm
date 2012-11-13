VERSION 5.00
Begin VB.Form frmQualify 
   BackColor       =   &H80000013&
   Caption         =   "Qualification"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10590
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToFrontPage 
      Caption         =   "Go back to Home Page"
      Height          =   615
      Left            =   10680
      TabIndex        =   33
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdExitQualify 
      Caption         =   "Exit Tax Program"
      Height          =   615
      Left            =   10680
      TabIndex        =   32
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue with Tax Return"
      Height          =   615
      Left            =   10680
      TabIndex        =   31
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm Qualification"
      Height          =   615
      Left            =   10680
      TabIndex        =   30
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmd7Details 
      BackColor       =   &H8000000C&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      MaskColor       =   &H00404040&
      TabIndex        =   29
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmd13Details 
      BackColor       =   &H8000000C&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      MaskColor       =   &H00404040&
      TabIndex        =   28
      Top             =   9840
      Width           =   855
   End
   Begin VB.CheckBox chk 
      Caption         =   "13."
      Height          =   495
      Index           =   13
      Left            =   9240
      TabIndex        =   27
      Top             =   9720
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "12."
      Height          =   495
      Index           =   12
      Left            =   9240
      TabIndex        =   25
      Top             =   9120
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "11."
      Height          =   495
      Index           =   11
      Left            =   9240
      TabIndex        =   23
      Top             =   8400
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "10."
      Height          =   495
      Index           =   10
      Left            =   9240
      TabIndex        =   21
      Top             =   7800
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "9."
      Height          =   495
      Index           =   9
      Left            =   9240
      TabIndex        =   19
      Top             =   7200
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "8."
      Height          =   495
      Index           =   8
      Left            =   9240
      TabIndex        =   17
      Top             =   6480
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "7."
      Height          =   495
      Index           =   7
      Left            =   9240
      TabIndex        =   15
      Top             =   5760
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "6."
      Height          =   495
      Index           =   6
      Left            =   9240
      TabIndex        =   13
      Top             =   5160
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "5."
      Height          =   495
      Index           =   5
      Left            =   9240
      TabIndex        =   11
      Top             =   4560
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "4."
      Height          =   495
      Index           =   4
      Left            =   9240
      TabIndex        =   9
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "3."
      Height          =   495
      Index           =   3
      Left            =   9240
      TabIndex        =   7
      Top             =   3240
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "2."
      Height          =   495
      Index           =   2
      Left            =   9240
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.CheckBox chk 
      Caption         =   "1."
      Height          =   495
      Index           =   1
      Left            =   9240
      TabIndex        =   2
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblMyName 
      Caption         =   "By Brent Mergen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   34
      Top             =   10200
      Width           =   1455
   End
   Begin VB.Label Q13 
      Caption         =   "13. Did you (or your spouse, if married) receive income from any of the following sources? (If no ""check""box)."
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Top             =   9840
      Width           =   7695
   End
   Begin VB.Label Q12 
      Caption         =   "12. Did you receive advance payments for the Earned Income Credit? If not ""check"" the box."
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   9240
      Width           =   8655
   End
   Begin VB.Label Q11 
      Caption         =   $"DoYouQualify.frx":0000
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   8520
      Width           =   8655
   End
   Begin VB.Label Q10 
      Caption         =   $"DoYouQualify.frx":0097
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   7920
      Width           =   8655
   End
   Begin VB.Label Q9 
      Caption         =   "9. Did you receive earnings from a Qualified State Tuition Program? (no = ""check"")."
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   7320
      Width           =   8655
   End
   Begin VB.Label Q8 
      Caption         =   $"DoYouQualify.frx":0151
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   6600
      Width           =   8655
   End
   Begin VB.Label Q7a 
      Caption         =   $"DoYouQualify.frx":022A
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   5880
      Width           =   8655
   End
   Begin VB.Label Q6 
      Caption         =   "6. Were you (or your spouse, if married) legally blind at the end of 2005? (Hopefully not, if no, be sure to ""check"")."
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   5280
      Width           =   8655
   End
   Begin VB.Label Q5 
      Caption         =   "5. Did you make any mortgage payments in 2005? (hopefully not, if no, be sure to ""check"" )."
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   4680
      Width           =   8655
   End
   Begin VB.Label Q4 
      Caption         =   "4. Were you (or your spouse, if married) 64 years of age or younger as of January 1, 2005?"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   4080
      Width           =   8655
   End
   Begin VB.Label Q3 
      Caption         =   $"DoYouQualify.frx":02BF
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   8655
   End
   Begin VB.Label QulifyInfo 
      BackColor       =   &H80000013&
      Caption         =   $"DoYouQualify.frx":0354
      Height          =   855
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   7695
   End
   Begin VB.Label Q2 
      Caption         =   "2. If you are married, do you and your spouse plan to file two separate returns? (If you're single, ""check"" this question.)"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   8655
   End
   Begin VB.Label Q1 
      Caption         =   "1. Was your taxable income less than $100,000? "
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   8655
   End
   Begin VB.Label QualifyingQuestionare 
      BackColor       =   &H80000013&
      Caption         =   "Do You Qualify For The 1040EZ?  Let's See!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "frmQualify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'Qualification (frmQualify)
'Brent Timothy Mergen
'24 March 2006
'Answer all of the questions and see if the 1040EZ is right for you!

Option Explicit


Private Sub cmd13Details_Click()
    MsgBox "1. Dividends from owning stocks" & Chr(13) & Chr(10) & "2. Gains,or losses, from the sale of stocks, bonds, or mutual funds" & Chr(13) & Chr(10) & "3. Self-employment income" & Chr(13) & Chr(10) & "4. Income from rentals, partnerships, S-corporations, or trusts" & Chr(13) & Chr(10) & "5. State tax refunds" & Chr(13) & Chr(10) & "6. Alimony, gambling winnings, or jury duty pay" & Chr(13) & Chr(10) & "7. Social Security benefits or income from retirement plans", , "Details of Question 13"
End Sub

Private Sub cmd7Details_Click()
    MsgBox "1. You have a loan but someone else has the primary obligation to repay it." & Chr(13) & Chr(10) & "2. Someone else claims you as a dependent.", , "Details of Question 7"
End Sub

Private Sub cmdConfirm_Click()
    Dim pos As Integer
    Dim Found As Boolean
    Found = False
    Do While Found = False And pos < 13
        pos = pos + 1
        If chk(pos).Value = 0 Then
            Found = True
        End If
    Loop
    If Found = False Then
        MsgBox "You meet all the qualifications, please proceed by clicking the 'Continue with Tax Return' button.", , "Success!"
    Else
        MsgBox "Please check question number " & pos, , "Failure!"
    End If
End Sub

Private Sub cmdContinue_Click()
    Dim pos As Integer
    Dim Found As Boolean
    Found = False
    If Found = False Then
        frmPersonalInfo.Visible = True 'shows next form
        frmQualify.Visible = False 'hides old form
    Else
        MsgBox "Please check question number " & pos, , "Failure!"
    End If
End Sub

Private Sub cmdExitQualify_Click()
    MsgBox "Sorry for the inconvenience, this tax form is not right for you.  Please contact Brent Mergen at btmergen@csbsju.edu with any tax questions.", , "Exit"
    End
End Sub

Private Sub cmdToFrontPage_Click()
    frmFrontpage.Visible = True 'shows next form
    frmQualify.Visible = False 'hides old form
End Sub
