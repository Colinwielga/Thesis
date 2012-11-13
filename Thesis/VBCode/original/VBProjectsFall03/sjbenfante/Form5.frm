VERSION 5.00
Begin VB.Form GiveMeMyDiscount 
   BackColor       =   &H00008000&
   Caption         =   "Form5"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   7365
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   8280
      TabIndex        =   6
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return To Previous Page"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   1215
      Left            =   840
      TabIndex        =   4
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtGroup 
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton cmdDiscount 
      Caption         =   "Tell Me What My Discount Is!"
      Height          =   1215
      Left            =   840
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox pbxResults 
      Height          =   1815
      Left            =   5520
      ScaleHeight     =   1755
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter Number Of People In Your Group Below"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "GiveMeMyDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer 'this says that i can only be an integer'

Private Sub cmdClear_Click()
pbxResults.Cls
'this will clear whatever is in the picture box'
End Sub

Private Sub cmdDiscount_Click()
pbxResults.Cls 'this will clear what is in the picture box before the user eneters in something else'
i = txtGroup.Text 'this is saying that i = what the user entered into the text box'
If i >= 14 Then
    pbxResults.Print "Please Call 1-800-427-PACK"
    pbxResults.Print "We Have Special Discounts For Groups Of Your Size"
    pbxResults.Print "Thank You"
    'these will all print out if the integer fits into this range'
ElseIf i >= 10 Then
    pbxResults.Print "Your Discount Is 2 Free Tickets"
    pbxResults.Print "You Also Get 2 Free Packer Dogs"
    pbxResults.Print "You Also Get 2 Free Lambeau Drinks"
    'these will all print out if the integer fits into this range'
ElseIf i >= 6 Then
    pbxResults.Print "Your Discount Is 1 Free Ticket"
    pbxResults.Print "You Also Get 1 Free Packer Dog"
    pbxResults.Print "You Also Get 1 Free Lambeau Drink"
    'these will all print out if the integer fits into this range'
ElseIf i >= 0 Then
    MsgBox "Sorry, But We Don't Give Away Discounts For The Size Of Your Group", , "Sorry"
    'this will pop up if the user enters an integer too large for the range'
End If

End Sub

Private Sub cmdQuit_Click()
    End
'this will automatically end the program'
End Sub

Private Sub cmdReturn_Click()
GiveMeMyDiscount.Hide
TicketPricing.Show
'this will hide the fifth form and show the second form'
End Sub

Private Sub Form_Load()
strPath = "n:\CS130\handin\sjbenfante\"
End Sub
