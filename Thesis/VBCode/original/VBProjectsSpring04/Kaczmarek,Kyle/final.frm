VERSION 5.00
Begin VB.Form frmfinal 
   BackColor       =   &H00400000&
   Caption         =   "Final"
   ClientHeight    =   7485
   ClientLeft      =   2370
   ClientTop       =   2370
   ClientWidth     =   10695
   FillColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10695
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdsh 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Include Shipping and Handling"
      Height          =   735
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdhats1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Change Hats"
      Height          =   975
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdcleats1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Change Cleats"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdpants1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Change Pants"
      Height          =   975
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdcjersey1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Change Jersey"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdfinalize 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Finalize Order"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   3855
   End
   Begin VB.PictureBox picresultsfinal 
      BackColor       =   &H8000000E&
      Height          =   6735
      Left            =   5880
      ScaleHeight     =   6675
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmfinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : BaseballUniforms (BaseballUniforms.vbp)
'Form Name : frmfinal (final.frm)
'Author: Kyle Kaczmarek
'Date Written: March 15, 2004
'Purpose of Form:'
Dim grandtotal As Single



Private Sub cmdcjersey1_Click()
frmjerseys.Show 'shows the jersey form
frmhats.Hide 'closes the hat form
frmorder.Hide 'closes the order form
frmpants.Hide 'closes the pants form
frmcleats.Hide 'closes the cleats form
frmfinal.Hide 'closes the final form
End Sub

Private Sub cmdcleats1_Click()
frmjerseys.Hide 'closes the jersey form
frmhats.Hide 'closes the hat form
frmorder.Hide 'closes the order form
frmpants.Hide 'closes the pants form
frmcleats.Show 'shows the cleats form
frmfinal.Hide 'closes the final form
End Sub

Private Sub cmdfinalize_Click()



picresultsfinal.Cls
picresultsfinal.Print "Item"; Tab(20); "Quantity"; Tab(35); "Cost" 'prints the titles
picresultsfinal.Print "*******"; Tab(20); "**********"; Tab(35); "*******" 'prints the stars
picresultsfinal.Print "Jerseys"; Tab(20); jerseys; Tab(35); FormatCurrency(jerseytotal, 2) 'prints the total for the jerseys
picresultsfinal.Print "Pants"; Tab(20); pants; Tab(35); FormatCurrency(pantstotal, 2) 'prints the total for the pants
picresultsfinal.Print "Hats"; Tab(20); hats; Tab(35); FormatCurrency(hatstotal, 2) 'prints the total for the hats
picresultsfinal.Print "Cleats"; Tab(20); cleats; Tab(35); FormatCurrency(cleatstotal, 2) 'prints the total for the cleats
picresultsfinal.Print Tab(35); "*******" 'prints the stars

grandtotal = jerseytotal + hatstotal + pantstotal + cleatstotal 'adds up the total for the jerseys, hats, pants, and cleats

picresultsfinal.Print "Your Total Is"; Tab(35); FormatCurrency(grandtotal, 2) 'prints the total for all 4 items

End Sub

Private Sub cmdhats1_Click()
frmjerseys.Hide 'closes the jersey form
frmhats.Show 'shows the hat form
frmorder.Hide 'closes the order form
frmpants.Hide 'closes the pants form
frmcleats.Hide 'closes the cleats form
frmfinal.Hide 'closes the final form
End Sub

Private Sub cmdpants1_Click()
frmjerseys.Hide 'closes the jersey form
frmhats.Hide 'closes the hat form
frmorder.Hide 'closes the order form
frmpants.Show 'closes the pants form
frmcleats.Hide 'closes the cleats form
frmfinal.Hide 'closes the final form
End Sub

Private Sub cmdquit_Click()
MsgBox "Thanks For Your Order!!!!!" 'a message box pops up
End 'ends the form
End Sub

Private Sub cmdsh_Click()

Dim number(1 To 10) As Integer, Shipping(1 To 10) As Single, Found As Boolean, POS As Integer, J As Integer

Found = False 'not found
POS = 0 'the position is 0
Open Path & "ShipandHand.txt" For Input As #1 'opens the file
For J = 1 To 10 'finds the number
    Input #1, number(J), Shipping(J) 'reads and stores the number
Next J 'moves to the next number
Do While Found = False And POS < 10
    POS = POS + 1 'moves to the position
    If grandtotal <= number(POS) Then 'compare the grandtotal with the number
        Found = True 'when it's found then it's true
        picresultsfinal.Print Tab(35); "*******" 'prints the stars
        SHfee = Shipping(POS) * grandtotal 'multiplies the shipping fee by the grand total
        picresultsfinal.Print "Shipping and Handling"; Tab(35); FormatCurrency(SHfee, 2) 'prints out the shipping shee
        picresultsfinal.Print Tab(35); "*******" 'prints out the stars
        newtotal = grandtotal + SHfee 'adds the new shipping fee and the grandtotal to find the final total
        picresultsfinal.Print "Final"; Tab(35); FormatCurrency(newtotal, 2) 'prints out the final total
    End If
Loop

Close #1 'closes the file

End Sub


Private Sub Form_Load()

Path = "N:\CS130\handin\Kaczmarek,Kyle\"

End Sub
