VERSION 5.00
Begin VB.Form frmParking 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   3615
      Left            =   9720
      ScaleHeight     =   3555
      ScaleWidth      =   3435
      TabIndex        =   19
      Top             =   4560
      Width           =   3495
   End
   Begin VB.CommandButton cmdPay2 
      Caption         =   "Pay"
      Height          =   735
      Left            =   2760
      TabIndex        =   18
      Top             =   6840
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   6000
      ScaleHeight     =   2715
      ScaleWidth      =   3315
      TabIndex        =   16
      Top             =   5400
      Width           =   3375
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   0
         Picture         =   "Parking.frx":0000
         ScaleHeight     =   6315
         ScaleWidth      =   5115
         TabIndex        =   17
         Top             =   -120
         Width           =   5175
      End
   End
   Begin VB.PictureBox picResults 
      Height          =   3615
      Left            =   3840
      ScaleHeight     =   3555
      ScaleWidth      =   5475
      TabIndex        =   15
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox txtExp 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtCC 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Text            =   "Please Enter Credit Card Number"
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Discover"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Master Card"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Visa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Payment"
      Enabled         =   0   'False
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdClick 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click Here to Start"
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Repeated Offenders"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10080
      TabIndex        =   20
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Expiration (MM/YY)"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Height          =   1335
      Left            =   2400
      TabIndex        =   12
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Height          =   1335
      Left            =   360
      TabIndex        =   11
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   10
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Height          =   1335
      Left            =   720
      TabIndex        =   8
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select Credit Card Type"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Pay Parking Tickets Here"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu Return 
         Caption         =   "Return to Menu"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmParking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Crime Awareness Project
'frm License
'By: Alex Warzecha & Andrew Enzler
'Written: 3/18/2009

'This program loads a file into two arrays, names and ticketprices
'It will search for names and attach the corresponding ticket price to that name

Private Sub cmdClick_Click()
Dim CTR As Integer, Names(1 To 100) As String, Ticketprices(1 To 200) As Single, Name As String
Dim Found As Boolean, j As Integer, Ticket As Single
'Initialize Variables
'Initialize CTR at 0

CTR = 0

'Load file into the two arrays using a Do until end of file loop
picResults2.Cls
Open App.Path & "\park.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1 'Increment Counter
    Input #1, Names(CTR), Ticketprices(CTR)
    picResults2.Print Names(CTR)
Loop
Close #1
'File should be loaded into two arrays

'Use an input box function to ask user for name

Name = InputBox("Enter name to see how many parking tickets you have:")

Found = False
'sets found to false
For j = 1 To CTR
    If Names(j) = Name Then
    Found = True ' found is set to true is the name input by the user is the same
                    ' as the name in the j position
    Ticket = Ticketprices(j)
    End If
Next j

If Not Found Then
    MsgBox "You do not have any parking ticket fees to pay at this time."
Else
    cmdPay.Enabled = True 'Command is enabled only if user needs to pay ticket
    picResults.Print "You owe "; FormatCurrency(Ticket); " in parking fees."
    picResults.Print "   "
    picResults.Print "Click on payment button below to pay ticket" 'user can choose to pay for ticket if name is found

End If



End Sub


Private Sub cmdPay_Click()

Check1.Enabled = True 'check boxes are enabled if the user has to pay fee
Check2.Enabled = True
Check3.Enabled = True


MsgBox "Please Select Credit Card Type" 'will ask user for credit card type via check box
MsgBox "Input any 8 digit number as a credit card number and 5 digit expiration date."

End Sub

Private Sub cmdPay2_Click()
Dim Credit As String, Exp As String, Found As Boolean


Credit = txtCC.Text ' uses text box in order to get input from user
Exp = txtExp.Text

If Len(Credit) = 8 And Len(Exp) = 5 Then  'will only accept a 8 digit number for credit card and a 4 digit for expiration
    picResults.Print " "
    picResults.Print "Thank you! Your fines are negated. You now owe: $0.00"

Else
MsgBox "Please enter a valid credit card number/expiration date"
' if number input by user is greater than 8 characters for the credit card number
' and greater than 5 numbers (including "/") for expiration date, a msg box will ask
'user to enter a number again
                                                
End If

End Sub



'offers to quit
Private Sub quit_Click()
End
End Sub

Private Sub return_Click()
frmParking.Hide 'Will return to home form
frmHome.Show
End Sub

