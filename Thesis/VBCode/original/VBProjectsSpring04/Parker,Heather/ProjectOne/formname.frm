VERSION 5.00
Begin VB.Form formname 
   BackColor       =   &H80000007&
   Caption         =   "A Solid Reputation"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form2"
   Picture         =   "formname.frx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   0
      Picture         =   "formname.frx":0DF0
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   0
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   1320
      Picture         =   "formname.frx":1C2C
      ScaleHeight     =   675
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "Back To Home"
      Height          =   855
      Left            =   7920
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picshow 
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   5400
      Width           =   3975
   End
   Begin VB.TextBox txtthing 
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdask 
      Caption         =   "Ask "
      Height          =   975
      Left            =   7440
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   4560
      Picture         =   "formname.frx":2209
      ScaleHeight     =   3195
      ScaleWidth      =   4995
      TabIndex        =   6
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   $"formname.frx":BF61
      Height          =   1335
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "formname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By:  Heather Parker
'With this form I tried many times to change the fonts and colors and each time
'the system froze on me and so I could not change them
'Form Name = formname
'Form File Name = Project One\formname.frm
'Purpose of the form:  Have th euser tell the prgram what fact
'they would like to hear more about concerning US Bank
'and have the program tell them the answer
Private Sub cmdask_Click()
Dim things(1 To 7) As String
Dim answer(1 To 7) As String
Dim J As Integer
M = txtthing
Open path & "usbank.txt" For Input As #2
Do While Not EOF(2)
    Ctr = Ctr + 1
    Input #2, things(Ctr), answer(Ctr)
Loop
Close #2
If M = things(1) Then
    picshow.Print "USBank is the "; answer(1); "Largest In the U.S."
End If
'Looks for items in a file that the user would like more informaton on and
'tells them when it is found
For J = 2 To Ctr
    If M = things(J) Then
        picshow.Print answer(J)
    End If
Next J

    
End Sub

Private Sub cmdhome_Click()
MsgBox "These are Only the Minimal things that USBank has done for their Name, To Learn More Log Onto www.USBank.org"
formname.Hide
Form1.Show
picshow.Cls
End Sub

Private Sub Form_Load()
Dim path As String
path = "N:\CS130\Parker, Heather\Project One\"
End Sub
