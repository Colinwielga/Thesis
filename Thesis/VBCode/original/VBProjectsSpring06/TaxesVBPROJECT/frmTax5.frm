VERSION 5.00
Begin VB.Form frmTax5 
   BackColor       =   &H80000013&
   Caption         =   "Line 5 Calculation"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click to view information for previous page"
      Height          =   615
      Left            =   1800
      TabIndex        =   22
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton cmdsubmit2 
      Caption         =   "Submit"
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   4920
      Width           =   1095
   End
   Begin VB.PictureBox picE 
      Height          =   255
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   20
      Top             =   5280
      Width           =   1095
   End
   Begin VB.PictureBox picD 
      Height          =   255
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTax5.frx":0000
      Left            =   2040
      List            =   "frmTax5.frx":000A
      TabIndex        =   17
      Text            =   "Please select an option"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.PictureBox picC 
      Height          =   255
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picB 
      Height          =   255
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Click to transfer information"
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton OptionE3 
      Caption         =   "Filing jointly, one is claimed as a dependent"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   4320
      Width           =   2655
   End
   Begin VB.OptionButton OptionE2 
      Caption         =   "Filing jointly, both are claimed as dependents "
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3840
      Width           =   2655
   End
   Begin VB.OptionButton OptionE1 
      Caption         =   "Single"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Return To Tax Form"
      Height          =   855
      Left            =   4920
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   6960
      Width           =   1695
   End
   Begin VB.PictureBox picDisplayCalc 
      Height          =   1815
      Left            =   1560
      ScaleHeight     =   1755
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H80000013&
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
      Left            =   7560
      TabIndex        =   23
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "No, enters $800"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Yes, adds $250 to line A"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblE 
      Caption         =   "E.  Exemption amount.  If you are..."
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label lblD 
      Caption         =   "D.  Enters smaller of line C and B."
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label lblC 
      Caption         =   "C.  If single, $5,000; if Married Filing Jointly, $10,000"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lblB 
      Caption         =   "B. Is line A more than $550?"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblA 
      Caption         =   "A.  Amount From Line 1 on Tax Form"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "frmTax5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'Line 5 Calc (frmTax5)
'Brent Timothy Mergen
'24 March 2006
'Calculates the 5th line for the tax return with some user inputs

Option Explicit
Dim txt1adj As Single
Dim txt2adj As Single
Dim singlenum As Single
Dim secondtotal, firsttotal As Single


Private Sub cmdGoBack_Click()
    frmTaxInput.Show
    frmTax5.Hide
End Sub

Private Sub cmdReset_Click()
    picOutput.Cls 'clears picbox
    picB.Cls 'clears picbox
    picC.Cls 'clears picbox
    picD.Cls 'clears picbox
    picE.Cls 'clears picbox
    picDisplayCalc.Cls 'clears picbox
End Sub

Private Sub cmdsubmit_Click()
    picC.Cls 'clears picbox
    picD.Cls 'clears picbox
        Dim a As Single
        Dim b As Single
        
        a = 5000
        b = 10000
        If Combo1 = "Single" Then
            picC.Print a
            singlenum = 5000
        End If
        If Combo1 = "Married" Then
            picC.Print b
            singlenum = 10000
        End If
        
        If txt1adj > singlenum Then
            picD.Print singlenum
            secondtotal = singlenum
        End If
        If txt1adj < singlenum Then
            picD.Print txt1adj
            secondtotal = txt1adj
        End If
    End Sub

Private Sub cmdsubmit2_Click()
    picE.Cls 'clears picbox
    If OptionE1 Then
        picE.Print "0"
        firsttotal = 0
    End If
    If OptionE2 Then
        picE.Print "0"
        firsttotal = 0
    End If
    If OptionE3 Then
        picE.Print "3200"
        firsttotal = 3200
    End If
End Sub

Private Sub cmdTransfer_Click()
    picOutput.Cls 'clears picbox
    picB.Cls 'clears picbox
    picOutput.Print txt1
    txt2adj = 800
    If txt1 >= 550 Then
        txt1adj = txt1 + 250
        picB.Print (txt1adj)
    Else
        picB.Print (txt2adj)
    End If
    
End Sub


Private Sub Command1_Click()
    picDisplayCalc.Cls 'clears picbox
    picDisplayCalc.Print firsttotal
    picDisplayCalc.Print secondtotal
    overalltotal = firsttotal + secondtotal
    picDisplayCalc.Print
    picDisplayCalc.Print overalltotal
End Sub

