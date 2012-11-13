VERSION 5.00
Begin VB.Form WhatIsTheCost 
   BackColor       =   &H0000C000&
   Caption         =   "Form3"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form3"
   ScaleHeight     =   6540
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return To Previous Page"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   5400
      Width           =   2655
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Find The Ticket Price"
      Height          =   1215
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   6255
      Left            =   3960
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   6195
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   0
      Width           =   5415
   End
   Begin VB.PictureBox pbxResults 
      Height          =   1575
      Left            =   720
      ScaleHeight     =   1515
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtSection 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblEnterData 
      Alignment       =   2  'Center
      Caption         =   "Enter The Section You Wish To Find A Price For Below. (0 to 38) "
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "WhatIsTheCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
WhatIsTheCost.Hide
TicketPricing.Show
'this will hide the third form and show the second form'
End Sub


Private Sub cmdClear_Click()
pbxResults.Cls
'this will clear whatever is inside the picture box'
End Sub

Private Sub cmdCompute_Click()
Dim i As Integer 'this says that i can only be a integer'
i = txtSection.Text 'this says that i is egual to what is entered in the text box'
If i >= 39 Then
    MsgBox "Sorry, but you must enter a number from 0 to 38", , "Error"
    ElseIf i >= 31 Then
        pbxResults.Print "Single Ticket Price Is $150.00"
        Total = Total + 150 'this says that this is what will be printed if it fits this condition'
    ElseIf i >= 25 Then
        pbxResults.Print "Single Ticket Price Is $270.00"
        Total = Total + 270 'this says that this is what will be printed if it fits this condition'
    ElseIf i >= 15 Then
        pbxResults.Print "Single Ticket Price Is $325.00"
        Total = Total + 325 'this says that this is what will be printed if it fits this condition'
    ElseIf i >= 9 Then
        pbxResults.Print "Single Ticket Price Is $270.00"
        Total = Total + 270 'this says that this is what will be printed if it fits this condition'
    ElseIf i >= 0 Then
        pbxResults.Print "Single Ticket Price Is $150.00"
        Total = Total + 150 'this says that this is what will be printed if it fits this condition'
    End If


End Sub


Private Sub cmdQuit_Click()
    End
'this automatically end the program'
End Sub

Private Sub cmdTotal_Click()
pbxResults.Print "*******************"
pbxResults.Print "Total Ticket Price Is:"; FormatCurrency(Total)
'this will print out what the running total is along with the words in quotations'
'FormatCurrency puts a dollar sign in fron to the total number printed out'
End Sub



Private Sub Form_Load()
strPath = "n:\CS130\handin\sjbenfante\"
End Sub
